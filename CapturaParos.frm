VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaParos 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Paros"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "CapturaParos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda de Datos"
      Height          =   8055
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6735
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   11880
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
      Begin VB.CommandButton CmdSalPro 
         Height          =   615
         Left            =   10560
         Picture         =   "CapturaParos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Sale de Busqueda"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtBuscli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "digite datos para buscar"
         Top             =   720
         Width           =   5655
      End
      Begin VB.OptionButton OptProCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptProDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "CapturaParos.frx":24B4
      Picture         =   "CapturaParos.frx":28F6
      Style           =   1  'Graphical
      TabIndex        =   184
      ToolTipText     =   "Primer Registro"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "CapturaParos.frx":2E28
      Picture         =   "CapturaParos.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   183
      ToolTipText     =   "Registro Anterior"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   10800
      MouseIcon       =   "CapturaParos.frx":379C
      Picture         =   "CapturaParos.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   182
      ToolTipText     =   "Siguiente Registro"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11160
      MouseIcon       =   "CapturaParos.frx":4110
      Picture         =   "CapturaParos.frx":4552
      Style           =   1  'Graphical
      TabIndex        =   181
      ToolTipText     =   "Ultimo Registro"
      Top             =   7560
      Width           =   375
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   8295
      Left            =   120
      TabIndex        =   38
      ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   1058
      BackColor       =   33023
      ForeColor       =   16711680
      TabCaption(0)   =   "Encabezado"
      TabPicture(0)   =   "CapturaParos.frx":4A84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle Tiempo"
      TabPicture(1)   =   "CapturaParos.frx":535E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridDetalleParos"
      Tab(1).Control(1)=   "FrameDetalle"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Detalle Consumo De Materias Primas"
      TabPicture(2)   =   "CapturaParos.frx":5678
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DbGridDetalleParos2"
      Tab(2).Control(1)=   "FrameDetalle2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Detalle Produccion"
      TabPicture(3)   =   "CapturaParos.frx":5F52
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DBGridInformacion"
      Tab(3).Control(1)=   "DBGridDetalleParos3"
      Tab(3).Control(2)=   "FrameDetalle3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Detalle Empleados"
      TabPicture(4)   =   "CapturaParos.frx":626C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DbGridEmpleados"
      Tab(4).Control(1)=   "FrameDetalle4"
      Tab(4).ControlCount=   2
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado de Paros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   120
         TabIndex        =   148
         Top             =   720
         Width           =   11415
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Editar"
            Height          =   735
            Left            =   2040
            Picture         =   "CapturaParos.frx":6288
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   6600
            Width           =   1100
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "B&uscar"
            Height          =   735
            Left            =   6840
            Picture         =   "CapturaParos.frx":665F
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   6600
            Width           =   1100
         End
         Begin VB.CommandButton CmdSalida 
            Caption         =   "&Salida"
            Height          =   735
            Left            =   9240
            Picture         =   "CapturaParos.frx":6AE7
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   6600
            Width           =   1215
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&Borrar"
            Height          =   735
            Left            =   5640
            Picture         =   "CapturaParos.frx":7002
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   6600
            Width           =   1100
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   735
            Left            =   4440
            Picture         =   "CapturaParos.frx":75CA
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   6600
            Width           =   1100
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   735
            Left            =   3240
            Picture         =   "CapturaParos.frx":7B01
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   6600
            Width           =   1100
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   735
            Left            =   840
            Picture         =   "CapturaParos.frx":805D
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6600
            Width           =   1100
         End
         Begin VB.Frame FrameInstalacion 
            Enabled         =   0   'False
            Height          =   6255
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   11175
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   24
               Left            =   5400
               MaxLength       =   15
               TabIndex        =   15
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   23
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   23
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   4080
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   22
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   22
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   3720
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   21
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   21
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   3360
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   20
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   20
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   3000
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   19
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   19
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   2040
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   18
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   18
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   1680
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   17
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   17
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   1320
               Width           =   3015
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   16
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   16
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   960
               Width           =   3015
            End
            Begin VB.TextBox TxtEficiencia 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               DataField       =   "Eficiencia"
               DataSource      =   "DataEncabezadoParos"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   1155
               Left            =   8640
               Locked          =   -1  'True
               TabIndex        =   166
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   5040
               Width           =   2415
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   14
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   13
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   4920
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Alignment       =   1  'Right Justify
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
               Height          =   288
               Index           =   1
               Left            =   1080
               MaxLength       =   12
               TabIndex        =   11
               Top             =   4200
               Width           =   1335
            End
            Begin VB.TextBox TxtDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   0
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   11
               Left            =   1080
               TabIndex        =   5
               ToolTipText     =   "los minutos tienen que estar en equivalente a horas"
               Top             =   2040
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Alignment       =   1  'Right Justify
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
               Height          =   288
               Index           =   10
               Left            =   1080
               MaxLength       =   12
               TabIndex        =   10
               Top             =   3840
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1080
               MaxLength       =   2
               TabIndex        =   12
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   4560
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   288
               Index           =   5
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   5280
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   1080
               MaxLength       =   1
               TabIndex        =   2
               Top             =   960
               Width           =   1335
            End
            Begin MSMask.MaskEdBox MskTurFin 
               Height          =   285
               Left            =   1080
               TabIndex        =   4
               Top             =   1680
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskTurIni 
               Height          =   285
               Left            =   1080
               TabIndex        =   3
               Top             =   1320
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   3
               Left            =   1080
               TabIndex        =   9
               Top             =   3480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   2
               Left            =   1080
               TabIndex        =   8
               Top             =   3120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   1
               Left            =   1080
               TabIndex        =   7
               ToolTipText     =   "Producto No Conforme"
               Top             =   2760
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   0
               Left            =   1080
               TabIndex        =   6
               ToolTipText     =   "Producto Conforme"
               Top             =   2400
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFec 
               Height          =   285
               Left            =   1080
               TabIndex        =   1
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtEfiRep 
               Height          =   375
               Left            =   8640
               TabIndex        =   167
               TabStop         =   0   'False
               Top             =   4320
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,###,##0.0000"
               PromptChar      =   "_"
            End
            Begin VB.Label LblGru 
               AutoSize        =   -1  'True
               Caption         =   "Grupo"
               Height          =   195
               Left            =   4800
               TabIndex        =   185
               Top             =   240
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Mecanico De Linea"
               Height          =   195
               Index           =   22
               Left            =   3240
               TabIndex        =   180
               Top             =   3360
               Width           =   1995
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Insp. De Aseg. Cal."
               Height          =   195
               Index           =   21
               Left            =   3240
               TabIndex        =   179
               Top             =   3720
               Width           =   1965
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Supervisor De Turno"
               Height          =   195
               Index           =   20
               Left            =   3240
               TabIndex        =   178
               Top             =   4080
               Width           =   2070
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Del Operador"
               Height          =   195
               Index           =   19
               Left            =   3240
               TabIndex        =   177
               Top             =   3000
               Width           =   1545
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Grupo De Manufactura Que Recibe La Linea"
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
               Index           =   26
               Left            =   3240
               TabIndex        =   176
               Top             =   2640
               Width           =   3810
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Mecanico De Linea"
               Height          =   195
               Index           =   25
               Left            =   3240
               TabIndex        =   175
               Top             =   1320
               Width           =   1995
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Insp. De Aseg. Cal."
               Height          =   195
               Index           =   24
               Left            =   3240
               TabIndex        =   174
               Top             =   1680
               Width           =   1965
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Supervisor De Turno"
               Height          =   195
               Index           =   23
               Left            =   3240
               TabIndex        =   173
               Top             =   2040
               Width           =   2070
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Grupo De Manufactura Que Entrega La Linea"
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
               Index           =   18
               Left            =   3240
               TabIndex        =   172
               Top             =   600
               Width           =   3870
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Del Operador"
               Height          =   195
               Index           =   17
               Left            =   3240
               TabIndex        =   171
               Top             =   960
               Width           =   1545
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Usuario"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   170
               Top             =   5280
               Width           =   540
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Efic. Reporte"
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
               Left            =   8640
               TabIndex        =   169
               Top             =   4080
               Width           =   1140
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "% Eficiencia"
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
               Left            =   8640
               TabIndex        =   168
               Top             =   4800
               Width           =   1050
            End
            Begin VB.Label LblEquipo 
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
               TabIndex        =   165
               Top             =   4920
               Width           =   4935
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Equipo"
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   164
               Top             =   4920
               Width           =   495
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Vel. Teorica"
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   163
               Top             =   4200
               Width           =   855
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Vel. Real"
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   162
               Top             =   3840
               Width           =   645
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Horas Progra."
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   161
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Termina"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   160
               Top             =   1680
               Width           =   570
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Inicia"
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   159
               Top             =   1320
               Width           =   375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Turno"
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   158
               Top             =   960
               Width           =   420
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   157
               Top             =   600
               Width           =   450
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   156
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Linea"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   155
               Top             =   4560
               Width           =   390
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Desper."
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   154
               ToolTipText     =   "Desperdicio"
               Top             =   3480
               Width           =   555
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Proceso"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   153
               ToolTipText     =   "Proceso"
               Top             =   3120
               Width           =   585
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "PNC"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   152
               ToolTipText     =   "Producto No Conforme"
               Top             =   2760
               Width           =   330
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "PC"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   151
               ToolTipText     =   "Producto Conforme"
               Top             =   2400
               Width           =   210
            End
            Begin VB.Label LblLinea 
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
               TabIndex        =   150
               Top             =   4560
               Width           =   4935
            End
         End
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "Imprimir"
            Height          =   735
            Left            =   8040
            Picture         =   "CapturaParos.frx":83DA
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   6600
            Width           =   1100
         End
      End
      Begin MSDataGridLib.DataGrid DbGridDetalleParos 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   147
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   2160
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   10
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
            DataField       =   "Orden"
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
         BeginProperty Column02 
            DataField       =   "Inicio"
            Caption         =   "Inicio"
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
            DataField       =   "Final"
            Caption         =   "Final"
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
            DataField       =   "Minutos"
            Caption         =   "Minutos"
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
            DataField       =   "Paro"
            Caption         =   "Codigo"
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
         BeginProperty Column06 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
            DataField       =   "DescripcionParo"
            Caption         =   "Paro"
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
            DataField       =   "Descripcion"
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
         BeginProperty Column09 
            DataField       =   "Empleado"
            Caption         =   "Empleado"
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
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   269.858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3300.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   840.189
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle 
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   110
         Top             =   720
         Width           =   11445
         Begin VB.Frame FrameTotal 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Total De Minutos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   138
            Top             =   6120
            Width           =   9855
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   142
               TabStop         =   0   'False
               ToolTipText     =   "debe cuadrar con horas programadas"
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Total"
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
               Left            =   7800
               TabIndex        =   146
               Top             =   120
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Paros Tipo S"
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
               Left            =   2520
               TabIndex        =   145
               Top             =   120
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Paros Tipo N"
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
               Left            =   5160
               TabIndex        =   144
               Top             =   120
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
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
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   143
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   495
            Left            =   2520
            Picture         =   "CapturaParos.frx":8914
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   6720
            Visible         =   0   'False
            Width           =   1500
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
            Left            =   8760
            TabIndex        =   136
            Top             =   6720
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   495
            Left            =   7200
            Picture         =   "CapturaParos.frx":8E46
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   6720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5640
            Picture         =   "CapturaParos.frx":9378
            Style           =   1  'Graphical
            TabIndex        =   134
            Top             =   6720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4080
            Picture         =   "CapturaParos.frx":98AA
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   6720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   960
            Picture         =   "CapturaParos.frx":9DDC
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   6720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Frame FrameDetalleInstalacion 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   120
            TabIndex        =   111
            Top             =   120
            Width           =   11295
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   288
               Index           =   6
               Left            =   720
               MaxLength       =   15
               TabIndex        =   114
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   288
               Index           =   7
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   113
               TabStop         =   0   'False
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   4
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   119
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   8760
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   8
               Left            =   4320
               MaxLength       =   10
               TabIndex        =   115
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   600
               Width           =   1215
            End
            Begin MSMask.MaskEdBox MskParFin 
               DataField       =   "Final"
               Height          =   285
               Left            =   840
               TabIndex        =   117
               Top             =   960
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
            Begin MSMask.MaskEdBox MskParIni 
               Height          =   285
               Left            =   120
               TabIndex        =   116
               Top             =   960
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
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Grupo"
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
               Left            =   8400
               TabIndex        =   187
               Top             =   960
               Width           =   525
            End
            Begin VB.Label LblParGru 
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
               Left            =   9000
               TabIndex        =   186
               Top             =   960
               Width           =   2175
            End
            Begin VB.Label Label11 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2160
               TabIndex        =   131
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label12 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   130
               Top             =   240
               Width           =   525
            End
            Begin VB.Label LblFicha 
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
               Left            =   5640
               TabIndex        =   129
               Top             =   240
               Width           =   5535
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo"
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
               Left            =   3720
               TabIndex        =   128
               Top             =   960
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   1
               Left            =   3360
               TabIndex        =   127
               Top             =   600
               Width           =   840
            End
            Begin VB.Label LblTipo 
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
               Height          =   285
               Left            =   4320
               TabIndex        =   126
               Top             =   960
               Width           =   735
            End
            Begin VB.Label LblParo 
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
               Left            =   5160
               TabIndex        =   125
               Top             =   960
               Width           =   3135
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Paro"
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
               Left            =   2640
               TabIndex        =   124
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Minutos"
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
               Left            =   1560
               TabIndex        =   123
               Top             =   720
               Width           =   675
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Inicio"
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
               TabIndex        =   122
               Top             =   720
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Final"
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
               Left            =   840
               TabIndex        =   121
               Top             =   720
               Width           =   420
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
               Left            =   5640
               TabIndex        =   120
               Top             =   600
               Width           =   5535
            End
         End
      End
      Begin MSDataGridLib.DataGrid DbGridDetalleParos2 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   109
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   2520
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   6
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
         ColumnCount     =   9
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
            DataField       =   "Orden"
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
         BeginProperty Column02 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column04 
            DataField       =   "FichaTecnica"
            Caption         =   "Fichatecnica"
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
         BeginProperty Column06 
            DataField       =   "Tarima"
            Caption         =   "Bulto"
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
            DataField       =   "Desperdicio"
            Caption         =   "Desperdicio"
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
         BeginProperty Column08 
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3344.882
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   870.236
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle2 
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   78
         Top             =   720
         Width           =   11445
         Begin VB.Frame FrameDetalleMateriaPrima 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   79
            Top             =   120
            Width           =   11295
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   8280
               TabIndex        =   86
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   7560
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   5
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   85
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               BackColor       =   &H80000004&
               Height          =   288
               Index           =   3
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   94
               TabStop         =   0   'False
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               DataField       =   "Orden"
               Height          =   288
               Index           =   4
               Left            =   720
               MaxLength       =   15
               TabIndex        =   82
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtTexto2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   6
               Left            =   10200
               TabIndex        =   88
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox TxtSaldo 
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
               Height          =   288
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox TxtTexto2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   8
               Left            =   7920
               TabIndex        =   87
               Top             =   1200
               Width           =   972
            End
            Begin VB.TextBox TxtLinCon 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               MaxLength       =   2
               TabIndex        =   84
               Text            =   "77"
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox TxtConLin 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5160
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSMask.MaskEdBox MskFecCon 
               Height          =   285
               Left            =   120
               TabIndex        =   83
               Top             =   840
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Consumo"
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
               Left            =   9240
               TabIndex        =   108
               Top             =   1200
               Width           =   780
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "# Bulto"
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
               Left            =   8280
               TabIndex        =   107
               Top             =   600
               Width           =   630
            End
            Begin VB.Label Label6 
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
               Left            =   1680
               TabIndex        =   106
               Top             =   600
               Width           =   1230
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "U/M"
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
               Left            =   10440
               TabIndex        =   105
               Top             =   600
               Width           =   345
            End
            Begin VB.Label LblFicha2 
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
               Left            =   5160
               TabIndex        =   104
               Top             =   240
               Width           =   6015
            End
            Begin VB.Label Label13 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   103
               Top             =   240
               Width           =   525
            End
            Begin VB.Label Label14 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2160
               TabIndex        =   102
               Top             =   240
               Width           =   1230
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
               Left            =   10200
               TabIndex        =   101
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Saldo"
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
               Index           =   7
               Left            =   9240
               TabIndex        =   100
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Desperdicio"
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
               Left            =   6720
               TabIndex        =   99
               Top             =   1200
               Width           =   1095
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
               Left            =   120
               TabIndex        =   98
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label6 
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
               Index           =   9
               Left            =   1080
               TabIndex        =   97
               Top             =   600
               Width           =   480
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
               Left            =   3000
               TabIndex        =   96
               Top             =   840
               Width           =   5175
            End
         End
         Begin VB.CommandButton CmdBotones3 
            Caption         =   "A&gregar"
            Height          =   495
            Index           =   0
            Left            =   960
            Picture         =   "CapturaParos.frx":A30E
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones3 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   2
            Left            =   2880
            Picture         =   "CapturaParos.frx":A840
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones3 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   3
            Left            =   4800
            Picture         =   "CapturaParos.frx":AD72
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones3 
            Caption         =   "B&orrar"
            Height          =   495
            Index           =   4
            Left            =   6720
            Picture         =   "CapturaParos.frx":B2A4
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones3 
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
            Index           =   5
            Left            =   8640
            TabIndex        =   93
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
      End
      Begin MSDataGridLib.DataGrid DBGridInformacion 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   76
         Top             =   4320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Descrip"
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
            DataField       =   "Descripcion"
            Caption         =   "Pasada"
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
         BeginProperty Column03 
            DataField       =   "Requerido"
            Caption         =   "Requerido"
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
         BeginProperty Column04 
            DataField       =   "Entregado"
            Caption         =   "Entregado"
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
         BeginProperty Column05 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
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
            DataField       =   "Desperdicio"
            Caption         =   "Desperdicio"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DBGridDetalleParos3 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   77
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   1920
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16744576
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
         ColumnCount     =   7
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
            DataField       =   "Orden"
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
         BeginProperty Column02 
            DataField       =   "Pasada"
            Caption         =   "Pasada"
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
            DataField       =   "ProductoConforme"
            Caption         =   "PC"
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
         BeginProperty Column04 
            DataField       =   "ProductoLiberado"
            Caption         =   "Proceso"
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
         BeginProperty Column05 
            DataField       =   "ProductoNoConforme"
            Caption         =   "PNC"
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
            DataField       =   "Desperdicio"
            Caption         =   "Desperdicio"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle3 
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   51
         Top             =   720
         Width           =   11445
         Begin VB.CommandButton CmdBotones4 
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
            Index           =   5
            Left            =   8640
            TabIndex        =   75
            Top             =   6720
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.CommandButton CmdBotones4 
            Caption         =   "B&orrar"
            Height          =   495
            Index           =   4
            Left            =   6720
            Picture         =   "CapturaParos.frx":B7D6
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones4 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   3
            Left            =   4800
            Picture         =   "CapturaParos.frx":BD08
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones4 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   2
            Left            =   2880
            Picture         =   "CapturaParos.frx":C23A
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones4 
            Caption         =   "A&gregar"
            Height          =   495
            Index           =   0
            Left            =   960
            Picture         =   "CapturaParos.frx":C76C
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Frame FrameDetalleProduccion 
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   11295
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   288
               Index           =   9
               Left            =   840
               MaxLength       =   15
               TabIndex        =   56
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   12
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   55
               TabStop         =   0   'False
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   11
               Left            =   6960
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   288
               Index           =   13
               Left            =   9000
               MaxLength       =   10
               TabIndex        =   57
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox ChkLam 
               Caption         =   "Lam. x Unid."
               Height          =   255
               Left            =   840
               TabIndex        =   53
               Top             =   600
               Width           =   1335
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   4
               Left            =   3480
               TabIndex        =   58
               ToolTipText     =   "Cantidad en Unidades"
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   5
               Left            =   7560
               TabIndex        =   60
               ToolTipText     =   "Cantidad en Unidades"
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   6
               Left            =   9720
               TabIndex        =   61
               ToolTipText     =   "Cantidad en Unidades"
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProducto 
               Height          =   285
               Index           =   7
               Left            =   5640
               TabIndex        =   59
               ToolTipText     =   "Cantidad en Unidades"
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2160
               TabIndex        =   70
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label9 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   69
               Top             =   240
               Width           =   525
            End
            Begin VB.Label LblFicha3 
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
               Height          =   225
               Left            =   5160
               TabIndex        =   68
               Top             =   240
               Width           =   3015
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "PC"
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
               Left            =   3000
               TabIndex        =   67
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "PNC"
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
               Left            =   7080
               TabIndex        =   66
               Top             =   600
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Desp."
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
               Left            =   9000
               TabIndex        =   65
               Top             =   600
               Width           =   510
            End
            Begin VB.Label LblPasada 
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
               Left            =   9720
               TabIndex        =   64
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Pasada"
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
               Left            =   8280
               TabIndex        =   63
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Proceso"
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
               Left            =   4920
               TabIndex        =   62
               Top             =   600
               Width           =   705
            End
         End
      End
      Begin MSDataGridLib.DataGrid DbGridEmpleados 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   50
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   1680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            DataField       =   "Empleado"
            Caption         =   "Codigo"
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
            DataField       =   "Descripcion"
            Caption         =   "Empleado"
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
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5279.812
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle4 
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   11445
         Begin VB.Frame FrameDetalleEmpleados 
            Enabled         =   0   'False
            Height          =   735
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   11175
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   15
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   47
               ToolTipText     =   "doble click o signo '+' para ayuda"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto2 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   9
               Left            =   8760
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label LblEmp 
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
               Left            =   2400
               TabIndex        =   49
               Top             =   240
               Width           =   8655
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.CommandButton CmdBotones5 
            Caption         =   "A&gregar"
            Height          =   495
            Index           =   0
            Left            =   840
            Picture         =   "CapturaParos.frx":CC9E
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones5 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   1
            Left            =   2760
            Picture         =   "CapturaParos.frx":D1D0
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones5 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Index           =   2
            Left            =   4680
            Picture         =   "CapturaParos.frx":D702
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones5 
            Caption         =   "B&orrar"
            Height          =   495
            Index           =   3
            Left            =   6600
            Picture         =   "CapturaParos.frx":DC34
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   6720
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBotones5 
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
            Index           =   4
            Left            =   8520
            TabIndex        =   40
            Top             =   6720
            Visible         =   0   'False
            Width           =   1680
         End
      End
   End
End
Attribute VB_Name = "CapturaParos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VDocumento As Double
Dim VDocumentoDetalle As Double

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim Bandera5 As Boolean
Dim BanderaBotonesVisibles As Boolean
Dim BanderaBotonesVisibles2 As Boolean
Dim BanderaBotonesVisibles3 As Boolean
Dim BanderaBotonesVisibles4 As Boolean
Dim BanderaBotonesVisiblesEncabezado As Boolean
Dim Criteria As String
Dim VUltimaFecha As Date
Dim VFechaActual As Date
Dim VHorasProgramadasEnMinutos As Integer
Dim BEditarConsumos As Boolean
Dim BEditarDetalle As Boolean
Dim vtexto As String
Dim VFechaProduccion As Date
Dim VLineaProduccion As String

'-------------------------
Dim REncabezadoParos As New ADODB.Recordset
Dim RDetalleParos As New ADODB.Recordset
Dim RDetalleConsumos As New ADODB.Recordset
Dim RDetalleEmpleados As New ADODB.Recordset
Dim RDetalleProduccion As New ADODB.Recordset
Dim RInformacion As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBuscaUnidadMedida As New ADODB.Recordset
Dim RCapturaParos As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaEncabezado As New ADODB.Recordset
Dim RBuscaSaldo As New ADODB.Recordset
Dim RBuscaUnico As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaFicha2 As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaParo As New ADODB.Recordset
Dim RBuscamaximo As New ADODB.Recordset
Dim RLineaActiva As New ADODB.Recordset
Dim RTotalP As New ADODB.Recordset
Dim RTotalS As New ADODB.Recordset
Dim RTotalN As New ADODB.Recordset
Dim RBuscaTurno As New ADODB.Recordset
Dim RBuscaEmpleado As New ADODB.Recordset
Dim RBuscaPasada As New ADODB.Recordset
Dim RBuscaEquipos As New ADODB.Recordset
Dim RSumaPC As New ADODB.Recordset
Dim RSumaPNC As New ADODB.Recordset
Dim RSumaD As New ADODB.Recordset
Dim RSumaP As New ADODB.Recordset
Dim RBuscaDetalleProduccion As New ADODB.Recordset
Dim RBuscaFichaOrden As New ADODB.Recordset
Dim RBuscaFichaOrden2 As New ADODB.Recordset
Dim RBuscaFichaOrden3 As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaEntradasMateriaPrima As New ADODB.Recordset
Dim RBuscaInformacion As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset
Dim RTiempoProgramadoD As New ADODB.Recordset
Dim RBuscaParosNoAfectanD As New ADODB.Recordset
Dim RBuscaParosSiAfectanD As New ADODB.Recordset
Dim RBuscaProduccionD As New ADODB.Recordset
Dim RProduccion As New ADODB.Recordset
Dim RBuscaEmpleados As New ADODB.Recordset
Dim RBuscaDetalleEmpleados As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RTotalCF As New ADODB.Recordset
Dim RTotalMP As New ADODB.Recordset

Dim BLinea As Boolean
Dim BFicha As Boolean
Dim BFicha2 As Boolean
Dim BFicha3 As Boolean
Dim BParo As Boolean
Dim BMateriaPrima As Boolean
Dim BUnidadMedida As Boolean
Dim BEmpleado As Boolean
Dim BEmpleado2 As Boolean
Dim BPasada As Boolean
Dim BEquipos As Boolean
Dim BEditar As Boolean
Dim BGrupo As Boolean

Dim VUltimaOrden As String
Dim VUltimaFicha As String
Dim VUltimoParo As String
Dim VUltimoEmpleado As String
Dim VFichaTecnica2 As String
Dim VInicio As String

Dim VTurIni As Double
Dim VTurTer As Double
Dim VHoras As Double
Dim VHoras1 As Double
Dim VHoras2 As Double
Dim VMinutos As Double
Dim VMinutos1 As Double
Dim VMinutos2 As Double
Dim VTotalMinutos As Double
Dim VTotalMinutosCurrency As Currency

Dim VPesoPorCuerpo As Single
Dim VFichaTecnica As String
Dim VTarima As Long
Dim VCantidad As Single
Dim VDesperdicio As Single
Dim VTipoParo As String

Dim VTiempoProgramadoD As Currency
Dim VVelocidadTeoricaDia As Integer
Dim VVelocidadRealDia As Integer
Dim VParosND As Single
Dim VParosSD As Single
Dim VProduccionD As Integer
Dim VTiempoRealProducidoD As Currency
Dim VPCD As Long
Dim VPNCD As Long
Dim VPDD As Long
Dim VTotalProduccionD As Long
Dim VFactor1D As Double
Dim VFactor2D As Double
Dim VFactor3D As Double
Dim VFactor4D As Double
Dim VFactor5D As Double
Dim VEficienciaRealD As Double
Dim VTotalParoS As Single
Dim VTotalParoN As Single
Dim VTotalParoP As Single

Dim VTotalProductoConforme As Long
Dim VTotalProductoNoConforme As Long
Dim VTotalDesperdicio As Long
Dim VUnidadesxLamina As Integer
Dim VUnidadesxLamina2 As Integer
Dim VLinea As String
Dim VPasada As String
Dim VPC As Long
Dim VPNC As Long
Dim VD As Long
Dim VP As Long

Dim VMensaje As String
Dim VMensaje2 As Long
Dim VContadorLinea As Integer

Dim VTotalParoCF As Integer
Dim VTotalParoMP As Integer



Sub Botones1()
    If Bandera = True Then
         FrameInstalacion.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         FrameDetalle.Visible = False
         DbGridDetalleParos.Visible = False
         FrameDetalle2.Visible = False
         DbGridDetalleParos2.Visible = False
         
         LblGru.Visible = True
         TxtTexto.Item(24).Visible = True
    Else
         FrameInstalacion.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         FrameDetalle.Visible = True
         DbGridDetalleParos.Visible = True
         FrameDetalle2.Visible = True
         DbGridDetalleParos2.Visible = True
         LblGru.Visible = False
         TxtTexto.Item(24).Visible = False
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleInstalacion.Enabled = True
         CmdAgregar2.Enabled = False
         CmdEditar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
         
    Else
         FrameDetalleInstalacion.Enabled = False
         CmdAgregar2.Enabled = True
         CmdEditar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If

End Sub

Sub BotonesVisibles()
    If BanderaBotonesVisibles = True Then
        CmdAgregar2.Visible = True
        CmdEditar2.Visible = True
        CmdBorrar2.Visible = True
        CmdGrabar2.Visible = True
        CmdCancelar2.Visible = True
        CmdTerminar.Visible = True
        
    Else
        CmdAgregar2.Visible = False
        CmdEditar2.Visible = False
        CmdBorrar2.Visible = False
        CmdGrabar2.Visible = False
        CmdCancelar2.Visible = False
        CmdTerminar.Visible = False
        
    End If
End Sub

Private Sub ChkLam_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
    
    Limpia_CamposDetalle
            BEditarDetalle = False
            DbGridDetalleParos.Enabled = False
            Bandera2 = True
            Botones2
            TxtTexto2.Item(0).Text = VDocumento
            TxtTexto.Item(6).Text = VUltimaOrden
            TxtTexto.Item(8).Text = VUltimoEmpleado
            MskParIni.Text = VUltimoParo
            TxtTexto2.Item(4).Text = "P"
            MskParFin.Text = "00:00"
            MskParFin.SetFocus
    
    
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

            If TxtDoc.Text = "" Then
                MsgBox "Documento Esta Vacio", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            
            If GBorrarEficiencia = False Then
                   MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                   Exit Sub
            End If

            VDocumento = TxtDoc.Text

            VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If VMensaje = vbOK Then
                MousePointer = 11
                     'BORRA EL REGISTRO
                        Conexion.Execute "delete from EncabezadoCapturaParos where documento = " & VDocumento
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        Limpia_CamposEncabezado
                        Limpia_CamposDetalle
                        Limpia_CamposConsumos
                        Limpia_CamposProduccion
                        
                        REncabezadoParos.Requery
                        REncabezadoParos.MoveLast
                        VDocumento = TxtDoc.Text
                        Llena_CamposEncabezado
                        
                        Set RDetalleParos = New ADODB.Recordset
                        Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                        
                        
                        'LLENA EL GRID
                        Set DbGridDetalleParos.DataSource = RDetalleParos
        
                        'SELECCIONA TODOS LOS DETALLES DE LOS CONSUMOS DE EL DOCUMENTO
                        Set RDetalleConsumos = New ADODB.Recordset
                        Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & VDocumento & " And D.FichaTecnica = F.Esp_Tec")
                        'LLENA EL GRID
                        Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                                
                        'SELECCIONA TODOS LOS DETALLES DE LA PRODUCCION
                        Set RDetalleProduccion = New ADODB.Recordset
                        Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & VDocumento)
                        'LLENA EL GRID
                        Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                        
                        'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                        'Limpia_CamposEmpleados
                        'Set RDetalleEmpleados = New ADODB.Recordset
                        'Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtTexto2.Item(9).Text & " And DE.Empleado = E.Codigo")
                        'LLENA EL GRID
                        'Set DbGridEmpleados.DataSource = RDetalleEmpleados
                        'Llena_CamposEmpleados
                        
                MousePointer = 0
            End If
                
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next

            
            VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If VMensaje = vbOK Then
                'BORRA EL REGISTRO
                        Conexion.Execute "Delete From DetalleCapturaParos Where Documento = " & TxtTexto2.Item(0).Text & " And Inicio = '" & MskParIni.Text & "'"
                        
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                            Set RDetalleParos = New ADODB.Recordset
                            Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & TxtDoc.Text & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                            
                            
                            
                            
                            'LLENA EL GRID
                            Set DbGridDetalleParos.DataSource = RDetalleParos
                            RDetalleParos.MoveLast
                            Llena_CamposDetalle
                        
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        
                
            End If
  
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        REncabezadoParos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        REncabezadoParos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        REncabezadoParos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        REncabezadoParos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If REncabezadoParos.BOF Then
        REncabezadoParos.MoveFirst
    ElseIf REncabezadoParos.EOF Then
        REncabezadoParos.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_CamposEncabezado
    
            'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO
                                Set RDetalleParos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & TxtDoc.Text & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                                
                                
                                'LLENA EL GRID
                                Set DbGridDetalleParos.DataSource = RDetalleParos
                                Llena_CamposDetalle

                                'SUMA LOS MINUTOS TIPO S
                                Set RTotalS = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'S'")
                                End If
                                    If RTotalS.RecordCount > 0 Then
                                        If IsNull(RTotalS(0)) Then
                                            TxtTotal.Item(0).Text = 0
                                        Else
                                            TxtTotal.Item(0).Text = RTotalS(0)
                                        End If
                                    Else
                                        TxtTotal.Item(0).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO N
                                Set RTotalN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'N'")
                                End If
                                    If RTotalN.RecordCount > 0 Then
                                        If IsNull(RTotalN(0)) Then
                                            TxtTotal.Item(1).Text = 0
                                        Else
                                            TxtTotal.Item(1).Text = RTotalN(0)
                                        End If
                                    Else
                                        TxtTotal.Item(1).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO P
                                Set RTotalP = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'P'")
                                End If

                                    If RTotalP.RecordCount > 0 Then
                                        If IsNull(RTotalP(0)) Then
                                            TxtTotal.Item(2).Text = 0
                                        Else
                                            TxtTotal.Item(2).Text = RTotalP(0)
                                        End If

                                    Else
                                        TxtTotal.Item(2).Text = 0
                                    End If

                                    
                                    'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                                    TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))


                                    'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                                    VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                                    'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                                    If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                       TxtTotal.Item(3).BackColor = vbWhite
                                    Else
                                        TxtTotal.Item(3).BackColor = vbYellow
                                    End If

               
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO PARA EL CONSUMO DE MATERIAS PRIMAS
                                Set RDetalleConsumos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & TxtDoc.Text & " And D.FichaTecnica = F.Esp_Tec")
                                'LLENA EL GRID
                                Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                Llena_CamposConsumos
                                

                               'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO EN PRODUCCION
                                Set RDetalleProduccion = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & TxtDoc.Text)
                                'LLENA EL GRID
                                Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                Llena_CamposProduccion
                   
                                'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                                'Limpia_CamposEmpleados
                                'Set RDetalleEmpleados = New ADODB.Recordset
                                'Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtDoc.Text & " And DE.Empleado = E.Codigo")
                                'LLENA EL GRID
                                'Set DbGridEmpleados.DataSource = RDetalleEmpleados
                                'Llena_CamposEmpleados
            
                                'SUMA EL TOTAL PRODUCTO CONFORME
                                Set RSumaPC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPC, "Select Sum(ProductoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPC.RecordCount > 0 Then
                                        If IsNull(RSumaPC(0)) Then
                                            VPC = 0
                                        Else
                                            VPC = RSumaPC(0)
                                        End If
                                            'SI EL TOTAL DEL PC = AL TOTAL PC EN EL ENCABEZADO
                                            If VPC = MskProducto.Item(0) Then
                                                MskProducto.Item(0).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(0).BackColor = vbYellow
                                            End If
                                    Else
                                        VPC = 0
                                    End If

                                'SUMA EL TOTAL PRODUCTO NO CONFORME
                                Set RSumaPNC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPNC, "Select Sum(ProductoNoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPNC.RecordCount > 0 Then
                                        If IsNull(RSumaPNC(0)) Then
                                            VPNC = 0
                                        Else
                                            VPNC = RSumaPNC(0)
                                        End If
                                            'SI EL TOTAL DEL PNC = AL TOTAL PNC EN EL ENCABEZADO
                                            If VPNC = MskProducto.Item(1) Then
                                                MskProducto.Item(1).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(1).BackColor = vbYellow
                                            End If
                                    Else
                                        VPNC = 0
                                    End If
                                
                                'SUMA EL TOTAL DE DESPERDICIO
                                Set RSumaD = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaD, "Select Sum(Desperdicio) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaD.RecordCount > 0 Then
                                        If IsNull(RSumaD(0)) Then
                                            VD = 0
                                        Else
                                            VD = RSumaD(0)
                                        End If
                                            'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                                            If VD = MskProducto.Item(3) Then
                                                MskProducto.Item(3).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(3).BackColor = vbYellow
                                            End If
                                    Else
                                        VD = 0
                                    End If
        
MousePointer = 0

End Sub

Private Sub CmdBotones3_Click(Index As Integer)
On Error Resume Next
        'AGREGAR
        If Index = 0 Then
                            Limpia_CamposConsumos
                    
                            DbGridDetalleParos2.Enabled = False
                            Bandera3 = True
                            Botones3
                            TxtTexto.Item(4).Text = VUltimaOrden
                            TxtTexto2.Item(5).Text = VFichaTecnica2
                            'TxtTexto.Item(3).Text = VUltimaFicha
                            TxtTexto2.Item(2).Text = VDocumento
                            'NUMERO DE INGRESO
                            TxtTexto2.Item(1).Text = 0
                            'CANTIDAD
                            TxtTexto2.Item(6).Text = 0
                            TxtTexto2.Item(8).Text = 0
                            TxtTexto.Item(4).SetFocus
                            
                            Set RBuscamaximo = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscamaximo, "Select Max(Documento) From DetalleConsumoMateriaPrima where Documento = " & VDocumento)
                                If RBuscamaximo.RecordCount > 0 Then
                                    If IsNull(RBuscamaximo(0)) Then
                                        TxtConLin.Text = "1"
                                    Else
                                        TxtConLin.Text = Val(RBuscamaximo(0)) + 1
                                    End If
                                Else
                                    TxtConLin.Text = "1"
                                End If
                    
                    
        'GRABAR
        ElseIf Index = 2 Then
                                                   
                        'VERIFICA ORDEN
                        If TxtTexto.Item(4) = "" Then
                            MsgBox "Orden De Producion No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(4).SetFocus
                            Exit Sub
                        End If
                        
                        'VERIFICA CODIGO DE MATERIA PRIMA
                        If TxtTexto2.Item(5) = "" Then
                            MsgBox "Codigo De Ficha Tecnica No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(5).SetFocus
                            Exit Sub
                        End If
                        
                        'NUMERO INGRESO
                        If Not IsNumeric(TxtTexto2.Item(1).Text) Then
                            MsgBox "Numero De Bulto Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(1).SetFocus
                            Exit Sub
                        End If
                                          
                        'DESPERDICIO
                        If Not IsNumeric(TxtTexto2.Item(8).Text) Then
                            MsgBox "Desperdicio Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(6).SetFocus
                            Exit Sub
                        End If
                        
                        'CANTIDAD CONSUMO
                        If Not IsNumeric(TxtTexto2.Item(6).Text) Then
                            MsgBox "Cantidad De Consumo Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(6).SetFocus
                            Exit Sub
                        End If
                        
                                         
                            
                        'GUARDA EN LA VARIABLE LA ULTIMA ORDEN DIGITADA
                        VUltimaOrden = TxtTexto.Item(4).Text
                        VFechaProduccion = MskFecCon.Text
                        VLineaProduccion = TxtLinCon.Text
                        VFichaTecnica2 = TxtTexto2.Item(5).Text
                        VFichaTecnica = TxtTexto2.Item(5).Text
                        VTarima = TxtTexto2.Item(1).Text
                        VCantidad = TxtTexto2.Item(6).Text
                        VDesperdicio = TxtTexto2.Item(8).Text
                        VContadorLinea = TxtConLin.Text
                        
                        
                        'REVISA SI EXISTE EL BULTO
                        Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima)
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = TO_DATE('" & VFechaProduccion & "','dd/mm/yyyy')" & " And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima)
                            End If
                            
                            If RBuscaTarima.RecordCount > 0 Then
                            Else
                                MsgBox "Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                            End If
                        
                           
                        'GRABA DATOS
                            
                            vtexto = VDocumento & ", '" 'DOCUMENTO
                            vtexto = vtexto & TxtTexto.Item(4).Text & "', " 'ORDEN
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 vtexto = vtexto & "#" & Format(MskFecCon.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                 vtexto = vtexto & "To_Date('" & Format(MskFecCon.Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            vtexto = vtexto & TxtLinCon.Text & "', '" 'LINEA
                            vtexto = vtexto & TxtTexto2.Item(5).Text & "', " 'FICHA TECNICA
                            vtexto = vtexto & TxtTexto2.Item(1).Text & ", " 'TARIMA
                            vtexto = vtexto & TxtTexto2.Item(8).Text & ", " 'DESPERDICIO
                            vtexto = vtexto & TxtTexto2.Item(6).Text & ", " 'CANTIDAD
                            vtexto = vtexto & VContadorLinea
                            
                            
                            'INICIA LA CONEXCION
                            Conexion.BeginTrans
                            
                                    'REALIZA EL INSERT
                                    Conexion.Execute "Insert Into DetalleConsumoMateriaPrima Values(" & vtexto & ")"
                    
                                        'SI SE DUPLICA LA LLAVE
                                         If GOrigenDeDatos = "AmaproAccess" Then
                                            If Err <> 0 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                Err.Clear
                                                Exit Sub
                                            End If
                                        Else 'ORACLE
                                            'sI ES CUALQUIER OTRO ERROR
                                            If Err = -2147217873 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Documento, Fecha, Linea, Ficha Tecnica, Tarima Ya Existen En Este Documento ", vbOKOnly + vbCritical, "Error"
                                                Err.Clear
                                                Exit Sub
                                            ElseIf Err <> 0 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                Err.Clear
                                                Exit Sub
                                            End If
                                        End If
                    
                        
                                    ' SALDO DEL BULTO ____________________________________________________
                                         If GOrigenDeDatos = "AmaproAccess" Then
                                                Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & (VCantidad + VDesperdicio) & " Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima
                                         Else
                                                
                                                Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & (VCantidad + VDesperdicio) & " Where FechaProduccion = TO_DATE('" & VFechaProduccion & "', 'dd/mm/yyyy') And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima
                                         End If
                                                                                             
                                         If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical & "Error"
                                            Exit Sub
                                         End If
                                         
                        'TERMINA LA CONEXION Y TERMINA DE GRABA LOS DATOS
                        Conexion.CommitTrans
                                    
                                    Bandera3 = False
                                    Botones3
                                    
                                    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
                                    Set RDetalleConsumos = New ADODB.Recordset
                                    Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & TxtDoc.Text & " And D.FichaTecnica = F.Esp_Tec")
                                    'LLENA EL GRID
                                    Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                    
                                    Llena_CamposConsumos
                                    
                                    DbGridDetalleParos2.Enabled = True
                                    CmdBotones3.Item(0).SetFocus
                                    
        
        'BORRAR
        ElseIf Index = 4 Then
                                If GBorrarEficiencia = False Then
                                       MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                                       Exit Sub
                                End If
                                
                                VDocumento = TxtTexto2.Item(2).Text
                                VFechaProduccion = MskFecCon.Text
                                VLineaProduccion = TxtLinCon.Text
                                VFichaTecnica = TxtTexto2.Item(5).Text
                                VTarima = TxtTexto2.Item(1).Text
                                VCantidad = TxtTexto2.Item(6).Text
                                VDesperdicio = TxtTexto2.Item(8).Text
                                VContadorLinea = TxtConLin.Text
                                
                                VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                    
                                'SI CONTESTA QUE SI QUIERE BORRAR
                                If VMensaje = vbOK Then
                                        MousePointer = 11
                                            
                                         'INICIA LA TRANSACCION
                                         Conexion.BeginTrans
                                         
                                                    'BORRA EL DETALLE
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Conexion.Execute "Delete From DetalleConsumoMateriaPrima Where Documento = " & VDocumento & " And Fecha = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And Contador = " & VContadorLinea
                                                    Else 'ORACLE
                                                        Conexion.Execute "Delete From DetalleConsumoMateriaPrima Where Documento = " & VDocumento & " And Fecha = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy') And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And Contador = " & VContadorLinea
                                                    End If
                                                    
                                                     
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            If Err <> 0 Then
                                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                Conexion.RollbackTrans
                                                                Err.Clear
                                                                Exit Sub
                                                            End If
                                                        Else 'ORACLE
                                                            'SI HAY ERRORES
                                                            If Err = -2147467259 Then
                                                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                                                Conexion.RollbackTrans
                                                                Err.Clear
                                                                Exit Sub
                                                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                Conexion.RollbackTrans
                                                                Err.Clear
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    
                                         
                                                     'BUSCA EL NUMERO DE INGRESO CON CODIGO DE MATERIA PRIMA EN EL DETALLE DE LAS ENTRADAS
                                                     'Y MODIFICA EL SALDO, SALIDAS Y PESO
                                                     'SALDO DEL BULTO ____________________________________________________
                                         
                                                     If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo + " & Val(VCantidad) + Val(VDesperdicio) & " Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima
                                                     Else
                                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo + " & Val(VCantidad) + Val(VDesperdicio) & " Where FechaProduccion = TO_DATE('" & VFechaProduccion & "', 'dd/mm/yyyy') And Linea = '" & VLineaProduccion & "' And FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima
                                                     End If
                                                    
                                                    If Err <> 0 Then
                                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical & "Error"
                                                        Conexion.RollbackTrans
                                                        Exit Sub
                                                    End If
                                        
                                            
                                        'TERMINA DE GRABAR
                                        Conexion.CommitTrans
                                            
                                            
                                        'SELECCIONA TODOS LOS DETALLES DE LA FACTURA
                                        Set RDetalleConsumos = New ADODB.Recordset
                                        Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & TxtDoc.Text & " And D.FichaTecnica = F.Esp_Tec")
                                        'LLENA EL GRID
                                        Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                        
                                        Llena_CamposConsumos
                                        
                                    
                                    MousePointer = 0
                                End If
                      
                                

        
        'CANCELAR
        ElseIf Index = 3 Then
                                
                                        Limpia_CamposConsumos
                                        
                                        DbGridDetalleParos2.Enabled = True
                                        Bandera3 = False
                                        Botones3
                                

        'TERMINAR
        ElseIf Index = 5 Then
                
                'HACE EL CALCULO DE EFICIENCIA
                CalculaEficiencia
                
                SumaParos

                
            'TERMINAR DE PAROS
            '-----------------------------------------------------------------------------------
                If CmdCancelar2.Enabled = True Then
                     CmdCancelar2_Click
                End If
                
                'DESHABILITA EL DETALLE Y HABILITA EL ENCABEZADO
                FrameDetalle.Visible = True
                FrameDetalle.Enabled = False
                FrameEncabezado.Enabled = True
                'BOTONES DE DATA
                CmdBotones2.Item(1).Visible = True
                CmdBotones2.Item(2).Visible = True
                CmdBotones2.Item(3).Visible = True
                CmdBotones2.Item(4).Visible = True

                
                'ESCONDE LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles = False
                BotonesVisibles
                
                'VISUALIZA LOS BOTONES DE ENCABEZADO
                BanderaBotonesVisiblesEncabezado = True
                BotonesVisiblesEncabezado
                
           'TERMINAR DE CONSUMOS
           '-----------------------------------------------------------------------------------
                If CmdBotones3.Item(3).Enabled = True Then
                     CmdBotones3_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle2.Visible = True
                FrameDetalle2.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles2 = False
                BotonesVisibles2
                
           'TERMINAR DE PRODUCCION
           '-----------------------------------------------------------------------------------
                If CmdBotones4.Item(3).Enabled = True Then
                     CmdBotones4_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle3.Visible = True
                FrameDetalle3.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles3 = False
                BotonesVisibles3
                
            'TERMINAR DE EMPLEADOS
                    '-----------------------------------------------------------------------------------
                         If CmdBotones5.Item(2).Enabled = True Then
                              CmdBotones5_Click (2)
                         End If
                    
                         'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                         FrameDetalle4.Visible = True
                         FrameDetalle4.Enabled = False
                         
                         'VISUALIZA LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles4 = False
                         BotonesVisibles4
                         
                         Llena_CamposEmpleados
                         
              TabDetalle.Tab = 0
                     
                    
        End If

        
End Sub

Private Sub CmdBotones4_Click(Index As Integer)
On Error Resume Next
        'AGREGAR
        If Index = 0 Then
                    Limpia_CamposProduccion
                    
                            DBGridDetalleParos3.Enabled = False
                            Bandera4 = True
                            Botones4
                            TxtTexto.Item(9).Text = VUltimaOrden
                            'TxtTexto.Item(12).Text = VUltimaFicha
                            TxtTexto2.Item(11).Text = VDocumento
                            TxtTexto.Item(9).SetFocus
                            MskProducto.Item(4).Text = 0
                            MskProducto.Item(5).Text = 0
                            MskProducto.Item(6).Text = 0
                            MskProducto.Item(7).Text = 0
                    
        'GRABAR
        ElseIf Index = 2 Then
                                                   
                     'REVISA LA ORDEN SI EXISTE
                        If TxtTexto.Item(9).Text <> "" Then
                           Set RBuscaOrden = New ADODB.Recordset
                           If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(9).Text & "'")
                           Else 'ORACLE
                                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtTexto.Item(9).Text) & "'")
                           End If
                               If RBuscaOrden.RecordCount > 0 Then
                               Else
                                  MsgBox "Numero De Orden No Existe", vbOKOnly + vbInformation, "Informacion"
                                  Exit Sub
                               End If
                        End If
                                                
                        'PRODUCTO CONFORME
                        If Not IsNumeric(MskProducto.Item(4).Text) Then
                            MsgBox "Producto Conforme Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(10).SetFocus
                            Exit Sub
                        End If
                                                
                        'PRODUCTO NO CONFORME
                        If Not IsNumeric(MskProducto.Item(5).Text) Then
                            MsgBox "Producto No Conforme Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(12).SetFocus
                            Exit Sub
                        End If
                        
                        'DESPERDICIO
                        If Not IsNumeric(MskProducto.Item(6).Text) Then
                            MsgBox "Desperdicio Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto2.Item(9).SetFocus
                            Exit Sub
                        End If
                        
                            
                        'GUARDA EN LA VARIABLE LA ULTIMA ORDEN DIGITADA
                        VUltimaOrden = TxtTexto.Item(9).Text
                        'GUARDA LA PASADA
                        VPasada = TxtTexto.Item(13).Text
                        'GUARDA EL TOTAL DE PRODUCTO CONFORME
                        VTotalProductoConforme = MskProducto.Item(4).Text
                        VTotalProductoNoConforme = MskProducto.Item(5).Text
                        VTotalDesperdicio = MskProducto.Item(6).Text
                        
                        'BUSCA EL DETALLE DE LA ORDEN DE LA PRODUCCION Y BUSCA LO REQUERIDO
                        Set RBuscaSaldo = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaSaldo, "Select * From DetalleOrdenProduccion Where Documento = '" & VUltimaOrden & "' And Linea = '" & VLinea & "' And Pasada = '" & VPasada & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaSaldo, "Select * From DetalleOrdenProduccion Where UPPER(Documento) = '" & UCase(VUltimaOrden) & "' And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(Pasada) = '" & UCase(VPasada) & "'")
                        End If
                            If RBuscaSaldo.RecordCount > 0 Then
                            Else
                                MsgBox "La Linea, Orden Y Pasada No Coinciden", vbOKOnly + vbInformation, "Revise"
                                Exit Sub
                            End If
                       
                            'GRABA DATOS
                            vtexto = VDocumento & ", '" 'DOCUMENTO
                            vtexto = vtexto & TxtTexto.Item(9).Text & "', " 'ORDEN
                            vtexto = vtexto & MskProducto.Item(4).Text & ", " 'PRODUCTO CONFORME
                            vtexto = vtexto & MskProducto.Item(7).Text & ", " 'PRODUCTO LIBERADO
                            vtexto = vtexto & MskProducto.Item(5).Text & ", " 'PRODUCTO NO CONFORME
                            vtexto = vtexto & MskProducto.Item(6).Text & ", '" 'DESPERDICIO
                            vtexto = vtexto & TxtTexto(13).Text & "'" 'PASADA
                            
                            'INICIA LA CONEXCION
                            Conexion.BeginTrans
                            
                                    'REALIZA EL INSERT
                                    Conexion.Execute "Insert Into DetalleProduccionPorOrden Values(" & vtexto & ")"
                                                                    
                                                                    
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                            If Err = -2147467259 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Oden y Pasada Ya Existen En Este Documento ", vbOKOnly + vbCritical, "Error"
                                                Exit Sub
                                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                Exit Sub
                                            End If
                                    Else
                                            If Err = -2147217873 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Oden y Pasada Ya Existen En Este Documento ", vbOKOnly + vbCritical, "Error"
                                                Exit Sub
                                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                Exit Sub
                                            End If
                                    End If
                                    
                                        'BUSCA EL DETALLE DE LA ORDEN Y ACTUALIZA
                                        'LOS SALDOS
                                        Conexion.Execute "Update DetalleOrdenProduccion Set Entregado = Entregado + (" & VTotalProductoConforme + VTotalProductoNoConforme & "), Saldo = Saldo - (" & VTotalProductoConforme + VTotalProductoNoConforme & "), Desperdicio = (Desperdicio + " & VTotalDesperdicio & ") Where Documento = '" & VUltimaOrden & "' And Linea = '" & VLinea & "' And Pasada = '" & VPasada & "'"
                                        
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                            Exit Sub
                                        End If
                                    
                            'GRABA LOS DATOS
                            Conexion.CommitTrans
                                                                                    
                                    Bandera4 = False
                                    Botones4
                                    
                                    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
                                    Set RDetalleProduccion = New ADODB.Recordset
                                    Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & VDocumento)
                                    'LLENA EL GRID
                                    Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                    
                                    DBGridDetalleParos3.Enabled = True
                                    CmdBotones4.Item(0).SetFocus
                                    
                                           'BUSCA LA INFORMACION DEL DETALLE DE LA ORDEN
                                            Set RInformacion = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RInformacion, "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo, DO.Desperdicio From DetalleOrdenProduccion DO, Lineas L, Pasadas P Where DO.Documento = '" & VUltimaOrden & "' And DO.Linea = L.Linea And DO.Pasada = P.Codigo")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RInformacion, "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo, DO.Desperdicio From DetalleOrdenProduccion DO, Lineas L, Pasadas P Where UPPER(DO.Documento) = '" & UCase(VUltimaOrden) & "' And DO.Linea = L.Linea And DO.Pasada = P.Codigo")
                                            End If
                                            'LLENA EL GRID
                                            Set DBGridInformacion.DataSource = RInformacion
                                                                                        
                                            DBGridInformacion.Columns(0).Width = "2000"
                                            DBGridInformacion.Columns(1).Width = "1500"
                                            DBGridInformacion.Columns(2).Width = "1200"
                                            DBGridInformacion.Columns(3).Width = "1200"
                                            DBGridInformacion.Columns(4).Width = "1200"
                                            DBGridInformacion.Columns(5).Width = "1200"
                                            DBGridInformacion.Columns(6).Width = "1200"
                                            DBGridInformacion.Columns(3).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(4).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(5).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(6).NumberFormat = "#,###,##0"
                                                              
                                    
                        
                        
                        
            'SUMA EL TOTAL PRODUCTO CONFORME
            Set RSumaPC = New ADODB.Recordset
            Call Abrir_Recordset(RSumaPC, "Select Sum(ProductoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                If RSumaPC.RecordCount > 0 Then
                    If IsNull(RSumaPC(0)) Then
                        VPC = 0
                    Else
                        VPC = RSumaPC(0)
                    End If
                        'SI EL TOTAL DEL PC = AL TOTAL PC EN EL ENCABEZADO
                        If VPC = MskProducto.Item(0) Then
                            MskProducto.Item(0).BackColor = vbWhite
                        Else
                            MskProducto.Item(0).BackColor = vbYellow
                        End If
                Else
                    VPC = 0
                End If
            
            'SUMA EL TOTAL PRODUCTO NO CONFORME
            Set RSumaPNC = New ADODB.Recordset
            Call Abrir_Recordset(RSumaPNC, "Select Sum(ProductoNoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                If RSumaPNC.RecordCount > 0 Then
                    If IsNull(RSumaPNC(0)) Then
                        VPNC = 0
                    Else
                        VPNC = RSumaPNC(0)
                    End If
                        'SI EL TOTAL DEL PNC = AL TOTAL PNC EN EL ENCABEZADO
                        If VPNC = MskProducto.Item(1) Then
                            MskProducto.Item(1).BackColor = vbWhite
                        Else
                            MskProducto.Item(1).BackColor = vbYellow
                        End If
                Else
                    VPNC = 0
                End If
            
            'SUMA EL TOTAL DE DESPERDICIO
            Set RSumaD = New ADODB.Recordset
            Call Abrir_Recordset(RSumaD, "Select Sum(Desperdicio) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                If RSumaD.RecordCount > 0 Then
                    If IsNull(RSumaD(0)) Then
                        VD = 0
                    Else
                        VD = RSumaD(0)
                    End If
                        'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                        If VD = MskProducto.Item(3) Then
                            MskProducto.Item(3).BackColor = vbWhite
                        Else
                            MskProducto.Item(3).BackColor = vbYellow
                        End If
                Else
                    VD = 0
                End If
                
                'SUMA EL TOTAL DE PROCESO
                Set RSumaP = New ADODB.Recordset
                    Call Abrir_Recordset(RSumaP, "Select Sum(ProductoLiberado) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                        If RSumaP.RecordCount > 0 Then
                            If IsNull(RSumaP(0)) Then
                                VP = 0
                            Else
                                VP = RSumaP(0)
                            End If
                        'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                            If VP = MskProducto.Item(2) Then
                                MskProducto.Item(2).BackColor = vbWhite
                            Else
                                MskProducto.Item(2).BackColor = vbYellow
                            End If
                        Else
                            VP = 0
                        End If
                    
        'BORRAR
        ElseIf Index = 4 Then
                                'documento
                                VDocumento = TxtTexto2.Item(11).Text
                                'GUARDA LA ORDEN
                                VUltimaOrden = TxtTexto.Item(9).Text
                                'GUARDA LA PASADA
                                VPasada = TxtTexto.Item(13).Text
                                'GUARDA EL TOTAL DE PRODUCTO CONFORME
                                VTotalProductoConforme = MskProducto.Item(4).Text
                                VTotalProductoNoConforme = MskProducto.Item(7).Text
                                VTotalDesperdicio = MskProducto.Item(6).Text
                        
                    
                                VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                    
                                'SI CONTESTA QUE SI QUIERE BORRAR
                                If VMensaje = vbOK Then
                                    MousePointer = 11
                                        
                                        'INICIA LA TRANSACCION
                                        Conexion.BeginTrans
                                        
                                                            'BORRA EL REGISTRO
                                                            Conexion.Execute "Delete From DetalleProduccionPorOrden Where Documento = " & VDocumento & " And Orden = '" & VUltimaOrden & "' And Pasada = '" & VPasada & "'"
                                                            
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    If Err <> 0 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    End If
                                                                Else 'ORACLE
                                                                    'SI HAY ERRORES
                                                                    If Err = -2147467259 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    ElseIf Err <> -2147467259 And Err <> 0 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    End If
                                                                End If
                                        
                                                            'BUSCA EL DETALLE DE LA ORDEN DE LA PRODUCCION Y BUSCA LO REQUERIDO
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Conexion.Execute "Update DetalleOrdenProduccion Set Entregado = Entregado - " & (VTotalProductoConforme + VTotalProductoNoConforme) & ", Saldo = Saldo + " & (VTotalProductoConforme + VTotalProductoNoConforme) & ", Desperdicio = Desperdicio - " & VTotalDesperdicio & " Where Documento = '" & VUltimaOrden & "' And Linea = '" & VLinea & "' And Pasada = '" & VPasada & "'"
                                                            Else 'ORACLE
                                                                Conexion.Execute "Update DetalleOrdenProduccion Set Entregado = Entregado - " & (VTotalProductoConforme + VTotalProductoNoConforme) & ", Saldo = Saldo + " & (VTotalProductoConforme + VTotalProductoNoConforme) & ", Desperdicio = Desperdicio - " & VTotalDesperdicio & " Where UPPER(Documento) = '" & UCase(VUltimaOrden) & "' And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(Pasada) = '" & UCase(VPasada) & "'"
                                                            End If
                                                    
                                                            If Err <> 0 Then
                                                                Conexion.RollbackTrans
                                                                MousePointer = 0
                                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                                Exit Sub
                                                            End If
                                                
                                        'GRABA LOS DATOS
                                        Conexion.CommitTrans
                                                                                
                                        'SELECCIONA TODOS LOS DETALLES DE LA FACTURA
                                        Set RDetalleProduccion = New ADODB.Recordset
                                        Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & VDocumento)
                                        'LLENA EL GRID
                                        Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                                                                
                                        If Err <> 0 Then
                                        End If
                                        
                                        Llena_CamposProduccion
                                        
                                    
                                    MousePointer = 0
                                End If
                      
                                
        
        'CANCELAR
        ElseIf Index = 3 Then
                                Limpia_CamposProduccion
                                
                                        DBGridDetalleParos3.Enabled = True
                                        Bandera4 = False
                                        Botones4
                                

        'TERMINAR
        ElseIf Index = 5 Then
                'HACE EL CALCULO DE EFICIENCIA
                CalculaEficiencia
                SumaParos
        
            'TERMINAR DE PAROS
            '-----------------------------------------------------------------------------------
                If CmdCancelar2.Enabled = True Then
                     CmdCancelar2_Click
                End If
                
                'DESHABILITA EL DETALLE Y HABILITA EL ENCABEZADO
                FrameDetalle.Visible = True
                FrameDetalle.Enabled = False
                FrameEncabezado.Enabled = True
                
                'BOTONES DE DATA
                CmdBotones2.Item(1).Visible = True
                CmdBotones2.Item(2).Visible = True
                CmdBotones2.Item(3).Visible = True
                CmdBotones2.Item(4).Visible = True
                
                'ESCONDE LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles = False
                BotonesVisibles
                
                'VISUALIZA LOS BOTONES DE ENCABEZADO
                BanderaBotonesVisiblesEncabezado = True
                BotonesVisiblesEncabezado
                
           'TERMINAR DE CONSUMOS
           '-----------------------------------------------------------------------------------
                If CmdBotones3.Item(3).Enabled = True Then
                     CmdBotones3_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle2.Visible = True
                FrameDetalle2.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles2 = False
                BotonesVisibles2
                    
           'TERMINAR DE PRODUCCION
           '-----------------------------------------------------------------------------------
                If CmdBotones4.Item(3).Enabled = True Then
                     CmdBotones4_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle3.Visible = True
                FrameDetalle3.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles3 = False
                BotonesVisibles3
                
                Llena_CamposProduccion
                
            'TERMINAR DE EMPLEADOS
           '-----------------------------------------------------------------------------------
                If CmdBotones5.Item(2).Enabled = True Then
                     CmdBotones5_Click (2)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle4.Visible = True
                FrameDetalle4.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles4 = False
                BotonesVisibles4
                
                Llena_CamposEmpleados

                TabDetalle.Tab = 0
        End If

        

End Sub

Private Sub CmdBotones5_Click(Index As Integer)
On Error Resume Next
        'AGREGAR
        If Index = 0 Then
            Limpia_CamposEmpleados
            DbGridEmpleados.Enabled = False
            Bandera5 = True
            Botones5
            TxtTexto2.Item(9).Text = VDocumento
            TxtTexto.Item(15).SetFocus
        
        
        'GRABAR
        ElseIf Index = 1 Then
        
                VDocumento = TxtTexto2.Item(9).Text
                    
                            'EMPLEADO
                            If TxtTexto.Item(15) = "" Then
                                MsgBox "Empleado No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(8).SetFocus
                                Exit Sub
                            End If
                                    
                            'GRABA DATOS
                                vtexto = VDocumento & ", '" 'DOCUMENTO
                                vtexto = vtexto & TxtTexto.Item(15) & "'" 'EMPLEADO
                                Conexion.Execute "Insert Into DetalleEmpleados Values(" & vtexto & ")"
                                                    
                            
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                        If Err = -2147467259 Then
                                            MousePointer = 0
                                            MsgBox "Documento y Empleado Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(15).SetFocus
                                            Exit Sub
                                      'SI ES CUALQUIER OTRO ERROR
                                        ElseIf Err <> -2147467259 And Err <> 0 Then
                                            MousePointer = 0
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            TxtTexto.Item(15).SetFocus
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                        If Err = -2147217873 Then
                                            MousePointer = 0
                                            MsgBox "Documento Y Empleado Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(15).SetFocus
                                            Exit Sub
                                      'SI ES CUALQUIER OTRO ERROR
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                            MousePointer = 0
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            TxtTexto.Item(15).SetFocus
                                            Exit Sub
                                        End If
                                    End If
                            
                            
                                        Bandera5 = False
                                        Botones5
                                        
                                        'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                                        Set RDetalleEmpleados = New ADODB.Recordset
                                        Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtTexto2.Item(9).Text & " And DE.Empleado = E.Codigo")
                                        
                                        'LLENA EL GRID
                                        Set DbGridEmpleados.DataSource = RDetalleEmpleados
                                        
                                        Llena_CamposEmpleados
                                        
                                        DbGridEmpleados.Enabled = True
                                        CmdBotones5.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 2 Then
                                        Limpia_CamposEmpleados
                                        DbGridEmpleados.Enabled = True
                                        Bandera5 = False
                                        Botones5
                         
        'BORRAR
        ElseIf Index = 3 Then
                                'documento
                                VDocumento = TxtTexto2.Item(9).Text
                                                    
                                VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                    
                                'SI CONTESTA QUE SI QUIERE BORRAR
                                If VMensaje = vbOK Then
                                    MousePointer = 11
                                                                                
                                                            'BORRA EL REGISTRO
                                                            Conexion.Execute "Delete From DetalleEmpleados Where Documento = " & VDocumento & " And Empleado = '" & TxtTexto.Item(15).Text & "'"
                                                            
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    If Err <> 0 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    End If
                                                                Else 'ORACLE
                                                                    'SI HAY ERRORES
                                                                    If Err = -2147467259 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    ElseIf Err <> -2147467259 And Err <> 0 Then
                                                                        Conexion.RollbackTrans
                                                                        MousePointer = 0
                                                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                        Err.Clear
                                                                    End If
                                                                End If
                                        
                                        
                                        'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                                        Set RDetalleEmpleados = New ADODB.Recordset
                                        Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & VDocumento & " And DE.Empleado = E.Codigo")
                                        
                                        'LLENA EL GRID
                                        Set DbGridEmpleados.DataSource = RDetalleEmpleados
                                                                                                                        
                                        If Err <> 0 Then
                                        End If
                                        
                                        Llena_CamposEmpleados
                                        
                                    
                                    MousePointer = 0
                                End If
        'TERMINAR
        ElseIf Index = 4 Then
                            'HACE EL CALCULO DE EFICIENCIA
                            CalculaEficiencia
                            SumaParos

        
                     'TERMINAR DE PAROS
                     '-----------------------------------------------------------------------------------
                         If CmdCancelar2.Enabled = True Then
                              CmdCancelar2_Click
                         End If
                         
                         'DESHABILITA EL DETALLE Y HABILITA EL ENCABEZADO
                         FrameDetalle.Visible = True
                         FrameDetalle.Enabled = False
                         FrameEncabezado.Enabled = True
                         
                         'BOTONES DE DATA
                         CmdBotones2.Item(1).Visible = True
                         CmdBotones2.Item(2).Visible = True
                         CmdBotones2.Item(3).Visible = True
                         CmdBotones2.Item(4).Visible = True
                         
                         'ESCONDE LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles = False
                         BotonesVisibles
                         
                         'VISUALIZA LOS BOTONES DE ENCABEZADO
                         BanderaBotonesVisiblesEncabezado = True
                         BotonesVisiblesEncabezado
                         
                    'TERMINAR DE CONSUMOS
                    '-----------------------------------------------------------------------------------
                         If CmdBotones3.Item(3).Enabled = True Then
                              CmdBotones3_Click (3)
                         End If
                    
                         'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                         FrameDetalle2.Visible = True
                         FrameDetalle2.Enabled = False
                         
                         'VISUALIZA LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles2 = False
                         BotonesVisibles2
                             
                    'TERMINAR DE PRODUCCION
                    '-----------------------------------------------------------------------------------
                         If CmdBotones4.Item(3).Enabled = True Then
                              CmdBotones4_Click (3)
                         End If
                    
                         'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                         FrameDetalle3.Visible = True
                         FrameDetalle3.Enabled = False
                         
                         'VISUALIZA LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles3 = False
                         BotonesVisibles3
                         
                         Llena_CamposProduccion
                         
                     'TERMINAR DE EMPLEADOS
                    '-----------------------------------------------------------------------------------
                         If CmdBotones5.Item(2).Enabled = True Then
                              CmdBotones5_Click (2)
                         End If
                    
                         'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                         FrameDetalle4.Visible = True
                         FrameDetalle4.Enabled = False
                         
                         'VISUALIZA LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles4 = False
                         BotonesVisibles4
                         
                         Llena_CamposEmpleados
                         
                         TabDetalle.Tab = 0
                      
        
        End If
                                                

End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
  VMensaje = InputBox("Documento a Buscar")
    If VMensaje <> "" Then
            VMensaje2 = VMensaje
            If IsNumeric(VMensaje2) Then
                
                Set RBuscaDocumento = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoCapturaParos Where Documento = " & VMensaje2)
                            If RBuscaDocumento.RecordCount > 0 Then
                            Else
                                MsgBox "Documento No Existe ", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                            End If
                    
                'si se esta buscando un numero menor se va al primer registro
                If Val(VMensaje2) < Val(TxtDoc.Text) Then
                    REncabezadoParos.MoveFirst
                End If
                    REncabezadoParos.Find "Documento = " & VMensaje
                
                                
                MousePointer = 11
                
                
                VDocumento = VMensaje
                
                Llena_CamposEncabezado
                
                
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO
                                Set RDetalleParos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                                
                                
                                'LLENA EL GRID
                                Set DbGridDetalleParos.DataSource = RDetalleParos
                                Llena_CamposDetalle


                                'SUMA LOS MINUTOS TIPO S
                                Set RTotalS = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'S'")
                                End If
                                    If RTotalS.RecordCount > 0 Then
                                        If IsNull(RTotalS(0)) Then
                                            TxtTotal.Item(0).Text = 0
                                        Else
                                            TxtTotal.Item(0).Text = RTotalS(0)
                                        End If
                                    Else
                                        TxtTotal.Item(0).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO N
                                Set RTotalN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'N'")
                                End If
                                    If RTotalN.RecordCount > 0 Then
                                        If IsNull(RTotalN(0)) Then
                                            TxtTotal.Item(1).Text = 0
                                        Else
                                            TxtTotal.Item(1).Text = RTotalN(0)
                                        End If
                                    Else
                                        TxtTotal.Item(1).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO P
                                Set RTotalP = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'P'")
                                End If

                                    If RTotalP.RecordCount > 0 Then
                                        If IsNull(RTotalP(0)) Then
                                            TxtTotal.Item(2).Text = 0
                                        Else
                                            TxtTotal.Item(2).Text = RTotalP(0)
                                        End If

                                    Else
                                        TxtTotal.Item(2).Text = 0
                                    End If

                                    
                                    'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                                    TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))


                                    'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                                    VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                                    'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                                    If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                       TxtTotal.Item(3).BackColor = vbWhite
                                    Else
                                        TxtTotal.Item(3).BackColor = vbYellow
                                    End If

               
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO PARA EL CONSUMO DE MATERIAS PRIMAS
                                Set RDetalleConsumos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & VDocumento & " And D.FichaTecnica = F.Esp_Tec")
                                'LLENA EL GRID
                                Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                Llena_CamposConsumos
                                

                               'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO EN PRODUCCION
                                Set RDetalleProduccion = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & VDocumento)
                                'LLENA EL GRID
                                Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                Llena_CamposProduccion
                   
                                'MEXICO AUN NO LO UTILIZA
                                'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                                'Limpia_CamposEmpleados
                                'Set RDetalleEmpleados = New ADODB.Recordset
                                'Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtDoc.Text & " And DE.Empleado = E.Codigo")
                                'LLENA EL GRID
                                'Set DbGridEmpleados.DataSource = RDetalleEmpleados
                                'Llena_CamposEmpleados
            
                                'SUMA EL TOTAL PRODUCTO CONFORME
                                Set RSumaPC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPC, "Select Sum(ProductoConforme) From DetalleProduccionPorOrden Where Documento = " & VDocumento)
                                    If RSumaPC.RecordCount > 0 Then
                                        If IsNull(RSumaPC(0)) Then
                                            VPC = 0
                                        Else
                                            VPC = RSumaPC(0)
                                        End If
                                            'SI EL TOTAL DEL PC = AL TOTAL PC EN EL ENCABEZADO
                                            If VPC = MskProducto.Item(0) Then
                                                MskProducto.Item(0).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(0).BackColor = vbYellow
                                            End If
                                    Else
                                        VPC = 0
                                    End If

                                'SUMA EL TOTAL PRODUCTO NO CONFORME
                                Set RSumaPNC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPNC, "Select Sum(ProductoNoConforme) From DetalleProduccionPorOrden Where Documento = " & VDocumento)
                                    If RSumaPNC.RecordCount > 0 Then
                                        If IsNull(RSumaPNC(0)) Then
                                            VPNC = 0
                                        Else
                                            VPNC = RSumaPNC(0)
                                        End If
                                            'SI EL TOTAL DEL PNC = AL TOTAL PNC EN EL ENCABEZADO
                                            If VPNC = MskProducto.Item(1) Then
                                                MskProducto.Item(1).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(1).BackColor = vbYellow
                                            End If
                                    Else
                                        VPNC = 0
                                    End If
                                
                                'SUMA EL TOTAL DE DESPERDICIO
                                Set RSumaD = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaD, "Select Sum(Desperdicio) From DetalleProduccionPorOrden Where Documento = " & VDocumento)
                                    If RSumaD.RecordCount > 0 Then
                                        If IsNull(RSumaD(0)) Then
                                            VD = 0
                                        Else
                                            VD = RSumaD(0)
                                        End If
                                            'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                                            If VD = MskProducto.Item(3) Then
                                                MskProducto.Item(3).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(3).BackColor = vbYellow
                                            End If
                                    Else
                                        VD = 0
                                    End If
        
                
                
                MousePointer = 0
            Else
                MsgBox "Solo Datos Numericos Se Aceptan", vbOKOnly + vbInformation, "Informacion"
            End If
    End If
            
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
            Llena_CamposEncabezado
                                
                'BUSCA EL DETALLE DEL ENCABEZADO
                If IsNumeric(TxtDoc.Text) Then
                
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO
                                Set RDetalleParos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & TxtDoc.Text & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                                
                                
                                
                                'LLENA EL GRID
                                Set DbGridDetalleParos.DataSource = RDetalleParos
                                Llena_CamposDetalle
                                

                                'SUMA LOS MINUTOS TIPO S
                                Set RTotalS = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'S'")
                                End If
                                    If RTotalS.RecordCount > 0 Then
                                        If IsNull(RTotalS(0)) Then
                                            TxtTotal.Item(0).Text = 0
                                        Else
                                            TxtTotal.Item(0).Text = RTotalS(0)
                                        End If
                                    Else
                                        TxtTotal.Item(0).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO N
                                Set RTotalN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'N'")
                                End If
                                    If RTotalN.RecordCount > 0 Then
                                        If IsNull(RTotalN(0)) Then
                                            TxtTotal.Item(1).Text = 0
                                        Else
                                            TxtTotal.Item(1).Text = RTotalN(0)
                                        End If
                                    Else
                                        TxtTotal.Item(1).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO P
                                Set RTotalP = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'P'")
                                End If

                                    If RTotalP.RecordCount > 0 Then
                                        If IsNull(RTotalP(0)) Then
                                            TxtTotal.Item(2).Text = 0
                                        Else
                                            TxtTotal.Item(2).Text = RTotalP(0)
                                        End If

                                    Else
                                        TxtTotal.Item(2).Text = 0
                                    End If

                                    
                                    'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                                    TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))


                                    'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                                    VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                                    'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                                    If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                       TxtTotal.Item(3).BackColor = vbWhite
                                    Else
                                        TxtTotal.Item(3).BackColor = vbYellow
                                    End If

               
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO PARA EL CONSUMO DE MATERIAS PRIMAS
                                Set RDetalleConsumos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & TxtDoc.Text & " And D.FichaTecnica = F.Esp_Tec")
                                'LLENA EL GRID
                                Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                Llena_CamposConsumos
                                

                               'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO EN PRODUCCION
                                Set RDetalleProduccion = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & TxtDoc.Text)
                                'LLENA EL GRID
                                Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                Llena_CamposProduccion
                   
            
                                'SUMA EL TOTAL PRODUCTO CONFORME
                                Set RSumaPC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPC, "Select Sum(ProductoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPC.RecordCount > 0 Then
                                        If IsNull(RSumaPC(0)) Then
                                            VPC = 0
                                        Else
                                            VPC = RSumaPC(0)
                                        End If
                                            'SI EL TOTAL DEL PC = AL TOTAL PC EN EL ENCABEZADO
                                            If VPC = MskProducto.Item(0) Then
                                                MskProducto.Item(0).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(0).BackColor = vbYellow
                                            End If
                                    Else
                                        VPC = 0
                                    End If

                                'SUMA EL TOTAL PRODUCTO NO CONFORME
                                Set RSumaPNC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPNC, "Select Sum(ProductoNoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPNC.RecordCount > 0 Then
                                        If IsNull(RSumaPNC(0)) Then
                                            VPNC = 0
                                        Else
                                            VPNC = RSumaPNC(0)
                                        End If
                                            'SI EL TOTAL DEL PNC = AL TOTAL PNC EN EL ENCABEZADO
                                            If VPNC = MskProducto.Item(1) Then
                                                MskProducto.Item(1).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(1).BackColor = vbYellow
                                            End If
                                    Else
                                        VPNC = 0
                                    End If
                                
                                'SUMA EL TOTAL DE DESPERDICIO
                                Set RSumaD = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaD, "Select Sum(Desperdicio) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaD.RecordCount > 0 Then
                                        If IsNull(RSumaD(0)) Then
                                            VD = 0
                                        Else
                                            VD = RSumaD(0)
                                        End If
                                            'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                                            If VD = MskProducto.Item(3) Then
                                                MskProducto.Item(3).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(3).BackColor = vbYellow
                                            End If
                                    Else
                                        VD = 0
                                    End If
        
                End If
            
            
            
            'ACTUALIZA EL TEXTO DE LA LINEA PARA QUE SIEMPRE ESTE DISPONIBLE DESPUES DE GRABAR
            TxtTexto.Item(0).Enabled = True
            
            'CAMBIA BOTONES
            Bandera = False
            Botones1
    

End Sub

Private Sub CmdCancelar2_Click()
    On Error Resume Next
            Llena_CamposDetalle
    
            DbGridDetalleParos.Enabled = True
            Bandera2 = False
            Botones2
    

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    
        
            'ASIGNAMOS A LA VARIABLE FECHA DEL SISTEMA MENOS 1
            VUltimaFecha = DateValue(Date) - 2
            VFechaActual = DateValue(Date)
                    
                    
            'SI PUEDE EDITAR NO VALIDA LAS FECHAS
            If GEditarEficiencia = True Then
            Else
                    If (DateValue(MskFec.Text) >= VUltimaFecha And DateValue(MskFec.Text) <= VFechaActual) Then
                    Else
                        MsgBox "No Puede EDITAR Reportes De 3 o mas dias de la fecha actual, Llame al Encargado", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
            End If
                    
            'VARIABLE PARA CONTROLAR SI ESTA EDITANDO
            BEditar = True
            TxtDoc.Enabled = False
            Bandera = True
            Botones1
            
            'SI PUEDE EDITAR DEJA MODIFICAR LA FECHA
             If GEditarEficiencia = True Then
                MskFec.Enabled = True
                TxtTexto.Item(0).Enabled = True
             Else
                MskFec.Enabled = False
                TxtTexto.Item(0).Enabled = False
             End If
             
             MskProducto.Item(0).SetFocus
             TxtTexto.Item(5).Text = GUsuario
    
    
End Sub

Private Sub CmdEditar2_Click()
On Error Resume Next
            VInicio = MskParIni.Text
            BEditarDetalle = True
            DbGridDetalleParos.Enabled = False
            Bandera2 = True
            Botones2
            MskParFin.SetFocus

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
       
    VDocumento = TxtTexto2.Item(0)
    
    
    'REVISA LA ORDEN SI EXISTE
    If TxtTexto.Item(6).Text <> "" Then
       Set RBuscaOrden = New ADODB.Recordset
       If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(6).Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtTexto.Item(6).Text) & "'")
        End If
           If RBuscaOrden.RecordCount > 0 Then
           Else
              MsgBox "Numero De Orden No Existe", vbOKOnly + vbInformation, "Informacion"
              Exit Sub
           End If
    End If
    
    'EMPLEADO
    If TxtTexto.Item(8) = "" Then
        MsgBox "Empleado No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtTexto.Item(8).SetFocus
        Exit Sub
    End If
    
    'VERIFICA CODIGO DE PARO
    If TxtTexto2.Item(4) = "" Then
        MsgBox "Codigo De Paro No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtTexto2.Item(4).SetFocus
        Exit Sub
    End If
        
    'GUARDA EN LA VARIABLE LA ULTIMA ORDEN DIGITADA
    VUltimaOrden = TxtTexto.Item(6).Text
    'GUARDA EN LA VARIABLE LA HORA FINAL DEL ULTIMO PARO
    VUltimoParo = MskParFin.Text
    VTipoParo = LblTipo.Caption
    'GUARDA EN LA VARIABLE EL ULTIMO EMPLEADO
    VUltimoEmpleado = TxtTexto.Item(8).Text
           
        
                'GRABA DATOS
                'AGREGAR
                If BEditarDetalle = False Then
                    vtexto = VDocumento & ", '" 'DOCUMENTO
                    vtexto = vtexto & TxtTexto.Item(6) & "', '" 'ORDEN
                    vtexto = vtexto & MskParIni.Text & "', '" 'INICIO
                    vtexto = vtexto & MskParFin.Text & "', " 'FINAL
                    vtexto = vtexto & TxtTexto2.Item(3).Text & ", '" 'MINUTOS
                    vtexto = vtexto & TxtTexto2.Item(4).Text & "', '" 'PARO
                    vtexto = vtexto & TxtTexto.Item(8) & "'" 'EMPLEADO
                    
                    Conexion.Execute "Insert Into DetalleCapturaParos Values(" & vtexto & ")"
                    
                Else 'EDITAR
                    vtexto = "Documento = " & VDocumento & ", " 'DOCUMENTO
                    vtexto = vtexto & "Orden = '" & TxtTexto.Item(6).Text & "', " 'ORDEN
                    vtexto = vtexto & "Inicio = '" & MskParIni.Text & "', " 'INICIO
                    vtexto = vtexto & "Final = '" & MskParFin.Text & "', " 'FINAL
                    vtexto = vtexto & "Minutos = '" & TxtTexto2.Item(3).Text & "', " 'MINUTOS
                    vtexto = vtexto & "Paro = '" & TxtTexto2.Item(4).Text & "', " 'PARO
                    vtexto = vtexto & "Empleado = '" & TxtTexto.Item(8).Text & "' " 'EMPLEADO
                    vtexto = vtexto & "Where Documento = " & VDocumento & " And Inicio = '" & VInicio & "'"
                    
                    Conexion.Execute "Update DetalleCapturaParos Set " & vtexto
                End If
        
                
                        'SI SE DUPLICA LA LLAVE
                         If GOrigenDeDatos = "AmaproAccess" Then
                            If Err = -2147467259 Then
                                MousePointer = 0
                                MsgBox "Documento y Campo De Inicio Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(0).SetFocus
                                Exit Sub
                          'SI ES CUALQUIER OTRO ERROR
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MousePointer = 0
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Exit Sub
                            End If
                        Else 'ORACLE
                            If Err = -2147217873 Then
                                MousePointer = 0
                                MsgBox "Documento Y Campo De Inicio Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(0).SetFocus
                                Exit Sub
                          'SI ES CUALQUIER OTRO ERROR
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                MousePointer = 0
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Exit Sub
                            End If
                        End If
                
                
                            Bandera2 = False
                            Botones2
                            
                            'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                            Set RDetalleParos = New ADODB.Recordset
                            Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & TxtDoc.Text & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                            
                            
                            'LLENA EL GRID
                            Set DbGridDetalleParos.DataSource = RDetalleParos
                            RDetalleParos.MoveLast
                            Llena_CamposDetalle
                            
                            DbGridDetalleParos.Enabled = True
                            CmdAgregar2.SetFocus
                            
                                    If VTipoParo = "S" Then
                                        'SUMA LOS MINUTOS TIPO S
                                        Set RTotalS = New ADODB.Recordset
                                        Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                                            If RTotalS.RecordCount > 0 Then
                                                If IsNull(RTotalS(0)) Then
                                                    TxtTotal.Item(0).Text = 0
                                                Else
                                                    TxtTotal.Item(0).Text = RTotalS(0)
                                                End If
                                            Else
                                                TxtTotal.Item(0).Text = 0
                                            End If
                                    End If
                                        
                                    If VTipoParo = "N" Then
                                        'SUMA LOS MINUTOS TIPO N
                                        Set RTotalN = New ADODB.Recordset
                                        Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                                            If RTotalN.RecordCount > 0 Then
                                                If IsNull(RTotalN(0)) Then
                                                    TxtTotal.Item(1).Text = 0
                                                Else
                                                    TxtTotal.Item(1).Text = RTotalN(0)
                                                End If
                                            Else
                                                TxtTotal.Item(1).Text = 0
                                            End If
                                    End If
                                    
                                    If VTipoParo = "P" Then
                                        'SUMA LOS MINUTOS TIPO P
                                        Set RTotalP = New ADODB.Recordset
                                        Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                                        
                                            If RTotalP.RecordCount > 0 Then
                                                If IsNull(RTotalP(0)) Then
                                                    TxtTotal.Item(2).Text = 0
                                                Else
                                                    TxtTotal.Item(2).Text = RTotalP(0)
                                                End If
                                                
                                            Else
                                                TxtTotal.Item(2).Text = 0
                                            End If
                                    End If
                                        
                                        'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                                            TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))
                                            
                            
                            'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                            VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                            'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                            If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                TxtTotal.Item(3).BackColor = vbWhite
                            Else
                                TxtTotal.Item(3).BackColor = vbYellow
                            End If
                                                        
    
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
   
        'VARIABLE PARA CONTROLAR SI ESTA EDITANDO
        BEditar = False
    
        Limpia_CamposEncabezado
    
        Bandera = True
        Botones1
        MskFec.Text = Format(Date, "dd/mm/yyyy")
        TxtTexto.Item(5).Text = GUsuario
        
        
        'BUSCA EL DOCUMENTO MAXIMO Y LE SUMA UN
        Set RBuscamaximo = New ADODB.Recordset
        Call Abrir_Recordset(RBuscamaximo, "Select Max(Documento) From EncabezadoCapturaParos")
            If RBuscamaximo.RecordCount > 0 Then
                If IsNull(RBuscamaximo(0)) Then
                    TxtDoc.Text = "1"
                Else
                    TxtDoc.Text = Val(RBuscamaximo(0)) + 1
                End If
            Else
                TxtDoc.Text = "1"
            End If
        
        TxtDoc.SetFocus
                
    
    
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
MousePointer = 11

    'CAMBIA EL FORMATO PARA EL AÑO CON 4 DIGISTOS PORQUE CON 2 DA PROBLEMA EN ORACLE
    If GOrigenDeDatos = "AmaproAccess" Then
    Else
        MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
    End If
    
    'VALIDA LA FECHA
    If Not IsDate(MskFec.Text) Then
            MousePointer = 0
            MsgBox "Fecha Incorrecta", vbOKOnly + vbCritical, "Error"
            MskFec.SetFocus
            Exit Sub
    End If
    
    'VALIDA EL CAMPO DE LINEA
    If TxtTexto.Item(0).Text = "" Then
            MousePointer = 0
            MsgBox "Linea No Puede Estar Vacia", vbOKOnly + vbCritical, "Informacion"
            TxtTexto.Item(0).SetFocus
            Exit Sub
    End If
    
    'PONE EN BLANCO LA VARIABLE DE ULTIMO EMPLEADO YA QUE A VECES SE LES OLVIDA
    'Y DEJAN EL CODIGO DEL EMPLEADO DEL REPORTE ANTERIOR
    VUltimoEmpleado = ""
    
    VDocumento = TxtDoc.Text
    VLinea = TxtTexto.Item(0).Text
    
    'VALIDA EL EQUIPO
    If TxtTexto.Item(14).Text = "" Then
        MousePointer = 0
        MsgBox "Equipo No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtTexto.Item(14).SetFocus
        Exit Sub
    End If
    
    'SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'REVISA SI YA EXISTE UN DOCUMENTO CON ESTA FECHA, LINEA Y TURNO
            Set RBuscaUnico = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaUnico, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# And Turno = '" & TxtTexto(2).Text & "' And Linea = '" & TxtTexto(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaUnico, "Select * From EncabezadoCapturaParos Where Fecha = TO_DATE('" & MskFec.Text & "', 'dd/mm/yyyy')" & " And Turno = '" & TxtTexto(2).Text & "' And Linea = '" & TxtTexto(0).Text & "'")
                    
                End If
                            If RBuscaUnico.RecordCount > 0 Then
                                MousePointer = 0
                                MsgBox "Esta Fecha, Linea y Turno, Ya Existen En Algun Documento ", vbOKOnly + vbInformation, "Informacion"
                                MskFec.SetFocus
                                Exit Sub
                            End If
    End If
    
    
                
            'GRABA DATOS
            If BEditar = False Then 'AGREGAR
                    vtexto = TxtDoc.Text & ", " 'DOCUMENTO
                    If GOrigenDeDatos = "AmaproAccess" Then
                        vtexto = vtexto & "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                    Else 'ORACLE
                        vtexto = vtexto & "TO_DATE('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                    End If
                    vtexto = vtexto & TxtTexto.Item(0).Text & "', '" 'LINEA
                    vtexto = vtexto & TxtTexto.Item(14).Text & "', '" 'GRUPO
                    vtexto = vtexto & TxtTexto.Item(2).Text & "', '" 'TURNO
                    vtexto = vtexto & MskTurIni.Text & "', '" 'INICIO
                    vtexto = vtexto & MskTurFin.Text & "', " 'FINAL
                    vtexto = vtexto & TxtTexto.Item(11).Text & ", "  'HORAS PROGRAMADAS
                    vtexto = vtexto & MskProducto.Item(0).Text & ", "  'PC
                    vtexto = vtexto & MskProducto.Item(1).Text & ", "  'PNC
                    vtexto = vtexto & MskProducto.Item(2).Text & ", "  'PPROCESO
                    vtexto = vtexto & MskProducto.Item(3).Text & ", "  'DESPERDICIO
                    vtexto = vtexto & TxtTexto.Item(1).Text & ", "  'VELOCIDAD TEORICA
                    vtexto = vtexto & TxtTexto.Item(10).Text & ", "  'VELOCIDAD REAL
                    vtexto = vtexto & TxtEfiRep.Text & ", "  'EFICIENCIA REPORTE
                    vtexto = vtexto & TxtEficiencia.Text & ", '"  'EFICIENCIA
                    vtexto = vtexto & TxtTexto.Item(5).Text & "',0,0,0, '"  'USUARIO
                    vtexto = vtexto & TxtTexto.Item(16).Text & "', '"  'OPERADOR ENTREGA
                    vtexto = vtexto & TxtTexto.Item(17).Text & "', '"  'MECANICO ENTREGA
                    vtexto = vtexto & TxtTexto.Item(18).Text & "', '"  'INSPECTOR ENTREGA
                    vtexto = vtexto & TxtTexto.Item(19).Text & "', '"  'SUPERVISOR ENTREGA
                    vtexto = vtexto & TxtTexto.Item(20).Text & "', '"  'OPERADOR RECIBE
                    vtexto = vtexto & TxtTexto.Item(21).Text & "', '"  'MECANICO RECIBE
                    vtexto = vtexto & TxtTexto.Item(22).Text & "', '"  'INSPECTOR RECIBE
                    vtexto = vtexto & TxtTexto.Item(23).Text & "', 0, 0"  'SUPERVISOR RECIBE
                    
                    
                    
                    'INSERTA EL DATO
                    Conexion.Execute "Insert Into EncabezadoCapturaParos Values(" & vtexto & ")"
            
            Else ' EDITAR
                    If GOrigenDeDatos = "AmaproAccess" Then
                        vtexto = "Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                    Else 'ORACLE
                        vtexto = "Fecha = TO_DATE('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                    End If
                    vtexto = vtexto & "Linea = '" & TxtTexto.Item(0).Text & "', " 'LINEA
                    vtexto = vtexto & "Grupo = '" & TxtTexto.Item(14).Text & "', " 'GRUPO
                    vtexto = vtexto & "Turno = '" & TxtTexto.Item(2).Text & "', " 'TURNO
                    vtexto = vtexto & "Inicio = '" & MskTurIni.Text & "', " 'INICIO
                    vtexto = vtexto & "Termina = '" & MskTurFin.Text & "', " 'FINAL
                    vtexto = vtexto & "HorasProgramadas = " & TxtTexto.Item(11).Text & ", "  'HORAS PROGRAMADAS
                    vtexto = vtexto & "ProductoConforme = " & MskProducto.Item(0).Text & ", "  'PC
                    vtexto = vtexto & "ProductoNoConforme = " & MskProducto.Item(1).Text & ", "  'PNC
                    vtexto = vtexto & "ProductoEnProceso = " & MskProducto.Item(2).Text & ", "  'PPROCESO
                    vtexto = vtexto & "Desperdicio = " & MskProducto.Item(3).Text & ", "  'DESPERDICIO
                    vtexto = vtexto & "VelocidadTeorica = " & TxtTexto.Item(1).Text & ", "  'VELOCIDAD TEORICA
                    vtexto = vtexto & "VelocidadReal = " & TxtTexto.Item(10).Text & ", "  'VELOCIDAD REAL
                    vtexto = vtexto & "EficienciaReporte = " & TxtEfiRep.Text & ", "  'EFICIENCIA REPORTE
                    vtexto = vtexto & "Eficiencia = " & TxtEficiencia.Text & ", "  'EFICIENCIA
                    vtexto = vtexto & "Usuario = '" & TxtTexto.Item(5).Text & "', "  'USUARIO
                    vtexto = vtexto & "OperadorEntrega = '" & TxtTexto.Item(16).Text & "', "  'OPERADOR ENTREGA
                    vtexto = vtexto & "MecanicoEntrega = '" & TxtTexto.Item(17).Text & "', "  'MECANICO ENTREGA
                    vtexto = vtexto & "InspectorEntrega = '" & TxtTexto.Item(18).Text & "', "  'INSPECTOR ENTREGA
                    vtexto = vtexto & "SupervisorEntrega = '" & TxtTexto.Item(19).Text & "', "  'SUPERVISOR ENTREGA
                    vtexto = vtexto & "OperadorRecibe = '" & TxtTexto.Item(20).Text & "', "  'OPERADOR RECIBE
                    vtexto = vtexto & "MecanicoRecibe = '" & TxtTexto.Item(21).Text & "', "  'MECANICO RECIBE
                    vtexto = vtexto & "InspectorRecibe = '" & TxtTexto.Item(22).Text & "', "  'INSPECTOR RECIBE
                    vtexto = vtexto & "SupervisorRecibe = '" & TxtTexto.Item(23).Text & "'"  'SUPERVISOR RECIBE
                    
                    vtexto = vtexto & "Where Documento = " & TxtDoc.Text
                    
                    'ACTUALIZA LOS DATOS EN BASE AL DOCUMENTO ACTUAL
                    Conexion.Execute "UPDATE EncabezadoCapturaParos SET " & vtexto
            
            End If
    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MousePointer = 0
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MousePointer = 0
                            MsgBox "Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MousePointer = 0
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                    
    
                    'CAMBIA BOTONES
                    Bandera = False
                    Botones1
                    
                    'ACTUALIZA EL TEXTO DE LA LINEA PARA QUE SIEMPRE ESTE DISPONIBLE DESPUES DE GRABAR
                     TxtTexto.Item(0).Enabled = True
            If BEditar = False Then 'AGREGAR
                     REncabezadoParos.Requery
                     REncabezadoParos.Find "Documento = " & VDocumento
            End If
                     'Set REncabezadoParos = New ADODB.Recordset
                     '       Call Abrir_Recordset(REncabezadoParos, "Select * From EncabezadoCapturaParos Where Documento = " & VDocumento)
                     '       If REncabezadoParos.RecordCount > 0 Then
                     '       Else
                     '       End If
                      
                     
                    
                    
                    Llena_CamposEncabezado
                                        
                    
                    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                    FrameDetalle.Enabled = True
                    FrameDetalle.Visible = True
                    FrameDetalle2.Enabled = True
                    FrameDetalle2.Visible = True
                    FrameDetalle3.Enabled = True
                    FrameDetalle3.Visible = True
                    FrameDetalle4.Enabled = True
                    FrameDetalle4.Visible = True
                    
    
                    FrameEncabezado.Enabled = False
                    
                    
                    
                    'MEXICO AUN NO LO USA, PORQUE NO TRABAJA EN EQUIPOS, ENTONCES NO SABESMOS QUE GENTE TRABAJA
                    
                    
                            'BUSCAMOS SI YA HAY EMPLEADOS EN ESTE DOCUMENTO
                                    'Set RBuscaDetalleEmpleados = New ADODB.Recordset
                                    '    Call Abrir_Recordset(RBuscaDetalleEmpleados, "Select * From DetalleEmpleados Where Documento = " & VDocumento)
                                    '        If RBuscaDetalleEmpleados.RecordCount > 0 Then
                                    '
                                    '        'SI NO HAY EMPLEADOS LOS AGREGA
                                    '        Else
                                    '                'BUSCA TODOS LOS EMPLEADOS QUE TIENE ESTE EQUIPO Y LOS AGREGA AL DETALLE
                                    '                Set RBuscaEmpleados = New ADODB.Recordset
                                    '                    Call Abrir_Recordset(RBuscaEmpleados, "Select Codigo From Empleados Where Grupo = '" & TxtTexto.Item(14).Text & "' And Estado = 'ALTA'")
                                    '
                                    '                    If RBuscaEmpleados.RecordCount > 0 Then
                                    '                            Do Until RBuscaEmpleados.EOF
                                    '                                    Conexion.Execute "insert into DetalleEmpleados Values(" & VDocumento & ", '" & RBuscaEmpleados!Codigo & "')"
                                    '
                                    '                                    If Err <> 0 Then
                                    '                                        MousePointer = 0
                                    '                                        MsgBox "Error En Agregar Los Empleados Del Equipo " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    '
                                    '                                   End If
                                    '                                RBuscaEmpleados.MoveNext
                                    '                            Loop
                                    '                    End If
                                    '        End If
                                    
    
                    'SELECCIONA TODOS LOS DETALLES DE PAROS DE EL DOCUMENTO
                    Limpia_CamposDetalle
                    Set RDetalleParos = New ADODB.Recordset
                    Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                    
                    
                    'LLENA EL GRID
                    Set DbGridDetalleParos.DataSource = RDetalleParos
                    Llena_CamposDetalle
    
                    'SELECCIONA TODOS LOS DETALLES DE LOS CONSUMOS DE EL DOCUMENTO
                    Limpia_CamposConsumos
                    Set RDetalleConsumos = New ADODB.Recordset
                    Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & VDocumento & " And D.FichaTecnica = F.Esp_Tec")
                    'LLENA EL GRID
                    Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                    Llena_CamposConsumos
                    
                    'SELECCIONA TODOS LOS DETALLES DE LA PRODUCCION
                    Limpia_CamposProduccion
                    Set RDetalleProduccion = New ADODB.Recordset
                    Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & VDocumento)
                    'LLENA EL GRID
                    Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                    Llena_CamposProduccion
                    
                    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                    'Limpia_CamposEmpleados
                    'Set RDetalleEmpleados = New ADODB.Recordset
                    'Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtDoc.Text & " And DE.Empleado = E.Codigo")
                    'LLENA EL GRID
                    'Set DbGridEmpleados.DataSource = RDetalleEmpleados
                    'Llena_CamposEmpleados
                                    
    
                    'VISUALIZA LOS BOTONES DEL DETALLE DE LA CAPTURA DE PAROS
                    BanderaBotonesVisibles = True
                    BotonesVisibles
                    
                    'VISUALIZA LOS BOTONES DEL DETALLE DE LOS CONSUMOS DE PAROS
                    BanderaBotonesVisibles2 = True
                    BotonesVisibles2
                    
                    'VISUALIZA LOS BOTONES DEL DETALLE PRODUCCION
                    BanderaBotonesVisibles3 = True
                    BotonesVisibles3
                    
                    'VISUALIZA LOS BOTONES DEL DETALLE EMPLEADOS
                    BanderaBotonesVisibles4 = True
                    BotonesVisibles4
                    
                    'SUMA LOS MINUTOS TIPO S
                        Set RTotalS = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                            If RTotalS.RecordCount > 0 Then
                                If IsNull(RTotalS(0)) Then
                                    TxtTotal.Item(0).Text = 0
                                Else
                                    TxtTotal.Item(0).Text = RTotalS(0)
                                End If
                            Else
                                TxtTotal.Item(0).Text = 0
                            End If
                        
                        'SUMA LOS MINUTOS TIPO N
                        Set RTotalN = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                            If RTotalN.RecordCount > 0 Then
                                If IsNull(RTotalN(0)) Then
                                    TxtTotal.Item(1).Text = 0
                                Else
                                    TxtTotal.Item(1).Text = RTotalN(0)
                                End If
                            Else
                                TxtTotal.Item(1).Text = 0
                            End If
                        
                        'SUMA LOS MINUTOS TIPO P
                        Set RTotalP = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                        
                            If RTotalP.RecordCount > 0 Then
                                If IsNull(RTotalP(0)) Then
                                    TxtTotal.Item(2).Text = 0
                                Else
                                    TxtTotal.Item(2).Text = RTotalP(0)
                                End If
                                
                            Else
                                TxtTotal.Item(2).Text = 0
                            End If
                            
                            'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                            TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))
                            
                            'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                            VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                            'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                            If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                TxtTotal.Item(3).BackColor = vbWhite
                            Else
                                TxtTotal.Item(3).BackColor = vbYellow
                            End If
                            
                            
                            
                            BanderaBotonesVisiblesEncabezado = False
                            BotonesVisiblesEncabezado
                            
                            'BOTONES DE DATA
                            CmdBotones2.Item(1).Visible = False
                            CmdBotones2.Item(2).Visible = False
                            CmdBotones2.Item(3).Visible = False
                            CmdBotones2.Item(4).Visible = False
                            
                            CmdAgregar2.SetFocus
                            TabDetalle.Tab = 1
                            
    MousePointer = 0

End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next
        
        MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "CapturaParos.rpt"
                Else
                    GNombreReporte = "CapturaParosO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoCapturaParos.Documento} = " & TxtDoc.Text
                FrmReporte.Show
            
        MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub CmdSalPro_Click()
    FrameBusqueda.Visible = False
    TxtBuscli.Text = ""
End Sub

Private Sub CmdTerminar_Click()
On Error Resume Next
            
            'HACE EL CALCULO DE EFICIENCIA
            CalculaEficiencia
            
            SumaParos

            'TERMINAR DE PAROS
            '-----------------------------------------------------------------------------------
                If CmdCancelar2.Enabled = True Then
                     CmdCancelar2_Click
                End If
                
                'DESHABILITA EL DETALLE Y HABILITA EL ENCABEZADO
                FrameDetalle.Visible = True
                FrameDetalle.Enabled = False
                FrameEncabezado.Enabled = True
                'BOTONES DE DATA
                CmdBotones2.Item(1).Visible = True
                CmdBotones2.Item(2).Visible = True
                CmdBotones2.Item(3).Visible = True
                CmdBotones2.Item(4).Visible = True
                
                'ESCONDE LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles = False
                BotonesVisibles
                
                'VISUALIZA LOS BOTONES DE ENCABEZADO
                BanderaBotonesVisiblesEncabezado = True
                BotonesVisiblesEncabezado
                
           'TERMINAR DE CONSUMOS
           '-----------------------------------------------------------------------------------
                If CmdBotones3.Item(3).Enabled = True Then
                     CmdBotones3_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle2.Visible = True
                FrameDetalle2.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles2 = False
                BotonesVisibles2
            
            'TERMINAR DE PRODUCCION
           '-----------------------------------------------------------------------------------
                If CmdBotones4.Item(3).Enabled = True Then
                     CmdBotones4_Click (3)
                End If
           
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                FrameDetalle3.Visible = True
                FrameDetalle3.Enabled = False
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                BanderaBotonesVisibles3 = False
                BotonesVisibles3
                
                'SE VUELVE A POSICIONAR EN EL DOCUMENTO PARA ACTUALIZAR EL CAMPO DE EFICIENCIA
                REncabezadoParos.Requery
                REncabezadoParos.Find ("Documento = " & VDocumento)
                Llena_CamposEncabezado
                Llena_CamposDetalle
                
                'TERMINAR DE EMPLEADOS
                    '-----------------------------------------------------------------------------------
                         If CmdBotones5.Item(2).Enabled = True Then
                              CmdBotones5_Click (2)
                         End If
                    
                         'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                         FrameDetalle4.Visible = True
                         FrameDetalle4.Enabled = False
                         
                         'VISUALIZA LOS BOTONES DEL DETALLE
                         BanderaBotonesVisibles4 = False
                         BotonesVisibles4
                         
                         Llena_CamposEmpleados
                     
                
                         TabDetalle.Tab = 0
                
                
End Sub




Private Sub DBGridBusqueda_DblClick()
        'LINEA EN ENCABEZADO DE PAROS
        If BLinea = True Then
            TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(0).SetFocus
        'FICHA TECNICA EN PAROS
        ElseIf BFicha = True Then
            TxtTexto.Item(7).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(7).SetFocus
        'CODIGO DE PARO
        ElseIf BParo = True Then
            TxtTexto2.Item(4).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(4).SetFocus
        'MATERIA PRIMA EN CONSUMOS
        ElseIf BMateriaPrima = True Then
            TxtTexto2.Item(5).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(5).SetFocus
        'FICHA TECNICA EN CONSUMOS
        ElseIf BFicha2 = True Then
            TxtTexto.Item(3).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(3).SetFocus
        'UNIDAD DE MEDIDA
        ElseIf BUnidadMedida = True Then
            TxtTexto2.Item(7).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(7).SetFocus
         'EMPLEADOS
        ElseIf BEmpleado = True Then
            TxtTexto.Item(8).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(8).SetFocus
        'FICHA TECNICA EN PRODUCCION
        ElseIf BFicha3 = True Then
            TxtTexto.Item(12).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(12).SetFocus
        'PASADAS
        ElseIf BPasada = True Then
            TxtTexto.Item(13).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(13).SetFocus
        'EQUIPOS
        ElseIf BEquipos = True Then
            TxtTexto.Item(14).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(14).SetFocus
        'EMPLEADOS 2
        ElseIf BEmpleado2 = True Then
            TxtTexto.Item(15).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(15).SetFocus
        'GRUPO
        ElseIf BGrupo = True Then
            TxtTexto.Item(24).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(24).SetFocus
        End If
            TxtBuscli.Text = ""
            FrameBusqueda.Visible = False
        

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        'LINEA EN ENCABEZADO DE PAROS
        If BLinea = True Then
            TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(0).SetFocus
        'FICHA TECNICA EN PAROS
        ElseIf BFicha = True Then
            TxtTexto.Item(7).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(7).SetFocus
        'CODIGO DE PARO
        ElseIf BParo = True Then
            TxtTexto2.Item(4).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(4).SetFocus
        'MATERIA PRIMA EN CONSUMOS
        ElseIf BMateriaPrima = True Then
            TxtTexto2.Item(5).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(5).SetFocus
        'FICHA TECNICA EN CONSUMOS
        ElseIf BFicha2 = True Then
            TxtTexto.Item(3).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(3).SetFocus
        'UNIDAD DE MEDIDA
        ElseIf BUnidadMedida = True Then
            TxtTexto2.Item(7).Text = DbGridBusqueda.Columns(0)
            TxtTexto2.Item(7).SetFocus
        'EMPLEADOS
        ElseIf BEmpleado = True Then
            TxtTexto.Item(8).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(8).SetFocus
        'FICHA TECNICA EN PRODUCCION
        ElseIf BFicha3 = True Then
            TxtTexto.Item(12).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(12).SetFocus
        'PASADAS
        ElseIf BPasada = True Then
            TxtTexto.Item(13).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(13).SetFocus
        'EQUIPOS
        ElseIf BEquipos = True Then
            TxtTexto.Item(14).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(14).SetFocus
        'EMPLEADOS 2
        ElseIf BEmpleado2 = True Then
            TxtTexto.Item(15).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(15).SetFocus
        'GRUPO
        ElseIf BGrupo = True Then
            TxtTexto.Item(24).Text = DbGridBusqueda.Columns(0)
            TxtTexto.Item(24).SetFocus
        End If
            TxtBuscli.Text = ""
            FrameBusqueda.Visible = False
    End If

End Sub




Private Sub DbGridDetalleParos_HeadClick(ByVal ColIndex As Integer)
        RDetalleParos.Sort = RDetalleParos.Fields(ColIndex).Name
End Sub


Private Sub DbGridDetalleParos_SelChange(Cancel As Integer)
    Llena_CamposDetalle
End Sub


Private Sub DbGridDetalleParos2_HeadClick(ByVal ColIndex As Integer)
        RDetalleConsumos.Sort = RDetalleConsumos.Fields(ColIndex).Name
End Sub


Private Sub DbGridDetalleParos2_SelChange(Cancel As Integer)
            Llena_CamposConsumos
End Sub


Private Sub DBGridDetalleParos3_HeadClick(ByVal ColIndex As Integer)
        RDetalleProduccion.Sort = RDetalleProduccion.Fields(ColIndex).Name
End Sub


Private Sub DBGridDetalleParos3_SelChange(Cancel As Integer)
            Llena_CamposProduccion
End Sub

Private Sub DbGridEmpleados_HeadClick(ByVal ColIndex As Integer)
        RDetalleEmpleados.Sort = RDetalleEmpleados.Fields(ColIndex).Name
End Sub

Private Sub DbGridEmpleados_SelChange(Cancel As Integer)
        Llena_CamposEmpleados
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
                
                MousePointer = 11
                
                Set REncabezadoParos = New ADODB.Recordset
                            Call Abrir_Recordset(REncabezadoParos, "Select * From EncabezadoCapturaParos Order by documento")
                            
                            REncabezadoParos.MoveLast
                                
                
                
                TxtDoc.Text = VMensaje
                VDocumento = VMensaje
                
                Llena_CamposEncabezado
                
                
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO
                                Set RDetalleParos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleParos, "Select DP.Documento, DP.Orden, DP.Inicio, DP.Final, DP.Minutos, DP.Paro, P.Tipo, P.DescripcionParo, PG.Descripcion, DP.Empleado from DetalleCapturaParos DP, Paros P, ParosGrupos PG where DP.Documento = " & TxtDoc.Text & " And DP.Paro = P.CodigoParo And P.Grupo = PG.CodigoGrupo Order By DP.Inicio")
                                
                                
                                'LLENA EL GRID
                                Set DbGridDetalleParos.DataSource = RDetalleParos
                                Llena_CamposDetalle
                                


                                'SUMA LOS MINUTOS TIPO S
                                Set RTotalS = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'S'")
                                End If
                                    If RTotalS.RecordCount > 0 Then
                                        If IsNull(RTotalS(0)) Then
                                            TxtTotal.Item(0).Text = 0
                                        Else
                                            TxtTotal.Item(0).Text = RTotalS(0)
                                        End If
                                    Else
                                        TxtTotal.Item(0).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO N
                                Set RTotalN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'N'")
                                End If
                                    If RTotalN.RecordCount > 0 Then
                                        If IsNull(RTotalN(0)) Then
                                            TxtTotal.Item(1).Text = 0
                                        Else
                                            TxtTotal.Item(1).Text = RTotalN(0)
                                        End If
                                    Else
                                        TxtTotal.Item(1).Text = 0
                                    End If

                                'SUMA LOS MINUTOS TIPO P
                                Set RTotalP = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & TxtDoc.Text & " And DC.Paro = P.CodigoParo And UPPER(P.Tipo) = 'P'")
                                End If

                                    If RTotalP.RecordCount > 0 Then
                                        If IsNull(RTotalP(0)) Then
                                            TxtTotal.Item(2).Text = 0
                                        Else
                                            TxtTotal.Item(2).Text = RTotalP(0)
                                        End If

                                    Else
                                        TxtTotal.Item(2).Text = 0
                                    End If

                                    
                                    'SUMA EL TOTAL DE LOS MINUTOS PAROS S, N Y PRODUCCION
                                    TxtTotal.Item(3).Text = Val(TxtTotal.Item(0)) + Val(TxtTotal.Item(1)) + Val(TxtTotal.Item(2))


                                    'CALCULA EN MINUTOS LAS HORAS PROGRAMADAS
                                    VHorasProgramadasEnMinutos = (TxtTexto.Item(11) * 60)
                                    'SI NO CUADRAN LAS HORAS PROGRAMADAS CON EL TOTAL DE MINUTOS DEL DETALLE
                                    If VHorasProgramadasEnMinutos = TxtTotal.Item(3).Text Then
                                       TxtTotal.Item(3).BackColor = vbWhite
                                    Else
                                        TxtTotal.Item(3).BackColor = vbYellow
                                    End If

               
                                'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO PARA EL CONSUMO DE MATERIAS PRIMAS
                                Set RDetalleConsumos = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleConsumos, "Select D.Documento, D.Orden, D.Fecha, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Desperdicio, D.Cantidad, D.Contador from DetalleConsumoMateriaPrima D, FichaTecnica F where D.Documento = " & TxtDoc.Text & " And D.FichaTecnica = F.Esp_Tec")
                                'LLENA EL GRID
                                Set DbGridDetalleParos2.DataSource = RDetalleConsumos
                                Llena_CamposConsumos
                                

                               'SELECCIONA TODOS LOS DETALLES DE EL DOCUMENTO EN PRODUCCION
                                Set RDetalleProduccion = New ADODB.Recordset
                                Call Abrir_Recordset(RDetalleProduccion, "Select * from DetalleProduccionPorOrden where Documento = " & TxtDoc.Text)
                                'LLENA EL GRID
                                Set DBGridDetalleParos3.DataSource = RDetalleProduccion
                                Llena_CamposProduccion
                   
                                'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL DOCUMENTO QUE SE ESTA GRABANDO
                                'Limpia_CamposEmpleados
                                'Set RDetalleEmpleados = New ADODB.Recordset
                                'Call Abrir_Recordset(RDetalleEmpleados, "Select DE.Documento, DE.Empleado, E.Descripcion from DetalleEmpleados DE, Empleados E where DE.Documento = " & TxtDoc.Text & " And DE.Empleado = E.Codigo")
                                'LLENA EL GRID
                                'Set DbGridEmpleados.DataSource = RDetalleEmpleados
                                'Llena_CamposEmpleados
            
                                'SUMA EL TOTAL PRODUCTO CONFORME
                                Set RSumaPC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPC, "Select Sum(ProductoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPC.RecordCount > 0 Then
                                        If IsNull(RSumaPC(0)) Then
                                            VPC = 0
                                        Else
                                            VPC = RSumaPC(0)
                                        End If
                                            'SI EL TOTAL DEL PC = AL TOTAL PC EN EL ENCABEZADO
                                            If VPC = MskProducto.Item(0) Then
                                                MskProducto.Item(0).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(0).BackColor = vbYellow
                                            End If
                                    Else
                                        VPC = 0
                                    End If

                                'SUMA EL TOTAL PRODUCTO NO CONFORME
                                Set RSumaPNC = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaPNC, "Select Sum(ProductoNoConforme) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaPNC.RecordCount > 0 Then
                                        If IsNull(RSumaPNC(0)) Then
                                            VPNC = 0
                                        Else
                                            VPNC = RSumaPNC(0)
                                        End If
                                            'SI EL TOTAL DEL PNC = AL TOTAL PNC EN EL ENCABEZADO
                                            If VPNC = MskProducto.Item(1) Then
                                                MskProducto.Item(1).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(1).BackColor = vbYellow
                                            End If
                                    Else
                                        VPNC = 0
                                    End If
                                
                                'SUMA EL TOTAL DE DESPERDICIO
                                Set RSumaD = New ADODB.Recordset
                                Call Abrir_Recordset(RSumaD, "Select Sum(Desperdicio) From DetalleProduccionPorOrden Where Documento = " & TxtDoc.Text)
                                    If RSumaD.RecordCount > 0 Then
                                        If IsNull(RSumaD(0)) Then
                                            VD = 0
                                        Else
                                            VD = RSumaD(0)
                                        End If
                                            'SI EL TOTAL DEL DESPERDICIO = AL TOTAL DE DESPERDICIO EN EL ENCABEZADO
                                            If VD = MskProducto.Item(3) Then
                                                MskProducto.Item(3).BackColor = vbWhite
                                            Else
                                                MskProducto.Item(3).BackColor = vbYellow
                                            End If
                                    Else
                                        VD = 0
                                    End If
                                 
                           MousePointer = 0
                
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

            REncabezadoParos.Close
            RDetalleParos.Close
            RDetalleConsumos.Close
            RDetalleProduccion.Close
            RInformacion.Close
            RBusqueda.Close
            RBuscaMateriaPrima.Close
            RBuscaUnidadMedida.Close
            
            RCapturaParos.Close
            RBuscaDetalle.Close
            RBuscaEncabezado.Close
            RBuscaSaldo.Close
            RBuscaUnico.Close
            RBuscaFicha.Close
            RBuscaLinea.Close
            RBuscaParo.Close
            RLineaActiva.Close
            RTotalP.Close
            RTotalS.Close
            RTotalN.Close
            RBuscaTurno.Close
            RBuscaEmpleado.Close
            RBuscaPasada.Close
            RBuscaEquipos.Close
            RSumaPC.Close
            RSumaPNC.Close
            RSumaD.Close
            RSumaP.Close
            RBuscaDetalleProduccion.Close
            RBuscaFichaOrden.Close
            RBuscaTarima.Close
            RBuscaEntradasMateriaPrima.Close
            
            RBuscaOrden.Close
            RTiempoProgramadoD.Close
            RBuscaParosNoAfectanD.Close
            RBuscaParosSiAfectanD.Close
            RBuscaProduccionD.Close
            RProduccion.Close

            Set REncabezadoParos = Nothing
            Set RDetalleParos = Nothing
            Set RDetalleConsumos = Nothing
            Set RDetalleProduccion = Nothing
            Set RInformacion = Nothing
            Set RBusqueda = Nothing
            Set RBuscaMateriaPrima = Nothing
            Set RBuscaUnidadMedida = Nothing
            
            Set RCapturaParos = Nothing
            Set RBuscaDetalle = Nothing
            Set RBuscaEncabezado = Nothing
            Set RBuscaSaldo = Nothing
            Set RBuscaUnico = Nothing
            Set RBuscaFicha = Nothing
            Set RBuscaLinea = Nothing
            Set RBuscaParo = Nothing
            Set RLineaActiva = Nothing
            Set RTotalP = Nothing
            Set RTotalS = Nothing
            Set RTotalN = Nothing
            Set RBuscaTurno = Nothing
            Set RBuscaEmpleado = Nothing
            Set RBuscaPasada = Nothing
            Set RBuscaEquipos = Nothing
            Set RSumaPC = Nothing
            Set RSumaPNC = Nothing
            Set RSumaD = Nothing
            Set RSumaP = Nothing
            Set RBuscaDetalleProduccion = Nothing
            Set RBuscaFichaOrden = Nothing
            Set RBuscaTarima = Nothing
            Set RBuscaEntradasMateriaPrima = Nothing
            
            Set RBuscaOrden = Nothing
            Set RTiempoProgramadoD = Nothing
            Set RBuscaParosNoAfectanD = Nothing
            Set RBuscaParosSiAfectanD = Nothing
            Set RBuscaProduccionD = Nothing
            Set RProduccion = Nothing
        If Err <> 0 Then
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

Private Sub MskFecCon_GotFocus()
        MskFecCon.SelStart = 0
        MskFecCon.SelLength = Len(MskFecCon.Text)
End Sub

Private Sub MskFecCon_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskParFin_GotFocus()
        MskParFin.SelStart = 0
        MskParFin.SelLength = Len(MskParFin.Text)
End Sub

Private Sub MskParFin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskParFin_LostFocus()
On Error Resume Next
        'DA ERROR A AL PONER 24 HORAS
        If MskParFin.Text = "24:00" Then
                    'RESTA LAS HORAS
                    VHoras = (Mid(MskParFin.Text, 1, 2) - Mid(MskParIni.Text, 1, 2))
                    'CONVIERTE LAS HORAS A MINUTOS
                    VHoras = VHoras * 60
                    'RESTA LOS MINUTOS
                    VMinutos = (Mid(MskParFin.Text, 4, 2) - Mid(MskParIni.Text, 4, 2))
                    'SUMA TODOS LOS MINUTOS
                    VTotalMinutos = VHoras + VMinutos
                                
                    TxtTexto2.Item(3).Text = VTotalMinutos
        ' SI NO ES 24 HORAS EL FINAL DEL PARO SIGUE NORMAL EL PROCESO
        Else
                'SI EL PARO FINAL YA PASO LAS 24 HORAS OSEA EL OTRO DIA
                If TimeValue(MskParIni.Text) > TimeValue(MskParFin.Text) Then
                    'RESTA LAS HORAS EN BASE A 24 HORAS
                    'HORAS ANTERIORES A LA MADRUGADA
                    VHoras1 = 24 - Mid(MskParIni.Text, 1, 2)
                    VHoras1 = VHoras1 - 1
                    
                    'HORAS DESPUES DE LA MADRUGADA
                    VHoras2 = Mid(MskParFin.Text, 1, 2)
                    'SUMA LOS DOS TIPOS DE HORAS CONVIERTE LAS HORAS A MINUTOS
                    VHoras = (VHoras1 + VHoras2) * 60
                    
                    'MINUTOS ANTERIORES A LA MADRUGADA
                    VMinutos1 = 60 - Mid(MskParIni.Text, 4, 2)
                    'MINUTOS DESPUES DE LA MADRUGADA
                    VMinutos2 = Mid(MskParFin.Text, 4, 2)
                    'SUMA LOS DOS TIPOS DE MINUTOS
                    VMinutos = (VMinutos1 + VMinutos2)
                    
                                        
                    'SUMA TODOS LOS MINUTOS
                    VTotalMinutos = VHoras + VMinutos
                                
                    TxtTexto2.Item(3).Text = VTotalMinutos
                Else
                    'RESTA LAS HORAS
                    VHoras = (Mid(MskParFin.Text, 1, 2) - Mid(MskParIni.Text, 1, 2))
                    'CONVIERTE LAS HORAS A MINUTOS
                    VHoras = VHoras * 60
                    'RESTA LOS MINUTOS
                    VMinutos = (Mid(MskParFin.Text, 4, 2) - Mid(MskParIni.Text, 4, 2))
                    'SUMA TODOS LOS MINUTOS
                    VTotalMinutos = VHoras + VMinutos
                                
                    TxtTexto2.Item(3).Text = VTotalMinutos
                End If
        End If
                    
End Sub

Private Sub MskParIni_GotFocus()
        MskParIni.SelStart = 0
        MskParIni.SelLength = Len(MskParIni.Text)
End Sub

Private Sub MskParIni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskProducto_GotFocus(Index As Integer)
        MskProducto.Item(Index).SelStart = 0
        MskProducto.Item(Index).SelLength = Len(MskProducto.Item(Index))
End Sub

Private Sub MskProducto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskProducto_LostFocus(Index As Integer)
        'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
        
                    Set RBuscaFicha = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(12).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(12).Text) & "'")
                    End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                        
                            If Index = 4 Then
                                If IsNumeric(MskProducto.Item(4).Text) Then
                                        MskProducto.Item(4).Text = MskProducto.Item(4).Text * VUnidadesxLamina
                                End If
                            ElseIf Index = 5 Then
                                If IsNumeric(MskProducto.Item(5).Text) Then
                                        MskProducto.Item(5).Text = MskProducto.Item(5).Text * VUnidadesxLamina
                                End If
                            ElseIf Index = 6 Then
                                If IsNumeric(MskProducto.Item(6).Text) Then
                                        MskProducto.Item(6).Text = MskProducto.Item(6).Text * VUnidadesxLamina
                                End If
                            ElseIf Index = 7 Then
                                If IsNumeric(MskProducto.Item(7).Text) Then
                                        MskProducto.Item(7).Text = MskProducto.Item(7).Text * VUnidadesxLamina
                                End If
                            End If
                        
        End If
End Sub

Private Sub MskTurFin_GotFocus()
        MskTurFin.SelStart = 0
        MskTurFin.SelLength = Len(MskTurFin.Text)
End Sub

Private Sub MskTurFin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskTurFin_LostFocus()
On Error Resume Next
        'DA ERROR A AL PONER 24 HORAS
       If MskTurFin.Text = "24:00" Then
                           'RESTA LAS HORAS
                    VHoras = (Mid(MskTurFin.Text, 1, 2) - Mid(MskTurIni.Text, 1, 2))
                    'RESTA LOS MINUTOS
                    VMinutos = (Mid(MskTurFin.Text, 4, 2) - Mid(MskTurIni.Text, 4, 2))
                    
                    If VMinutos > 0 Then
                        TxtTexto.Item(11).Text = VHoras & "." & VMinutos
                    Else
                        'SUMA TODOS LOS MINUTOS
                        TxtTexto.Item(11).Text = VHoras
                    End If
        ' SI NO ES 24 HORAS EL FINAL DEL PARO SIGUE NORMAL EL PROCESO
        Else
                'SI EL PARO FINAL YA PASO LAS 24 HORAS OSEA EL OTRO DIA
                If TimeValue(MskTurIni.Text) > TimeValue(MskTurFin.Text) Then
                    'RESTA LAS HORAS EN BASE A 24 HORAS
                    'HORAS ANTERIORES A LA MADRUGADA
                    VHoras1 = 24 - Mid(MskTurIni.Text, 1, 2)
                    VHoras1 = VHoras1 - 1
                    
                    'HORAS DESPUES DE LA MADRUGADA
                    VHoras2 = Mid(MskTurFin.Text, 1, 2)
                    'SUMA LOS DOS TIPOS DE HORAS CONVIERTE LAS HORAS A MINUTOS
                    VHoras = (VHoras1 + VHoras2)
                    
                    'MINUTOS ANTERIORES A LA MADRUGADA
                    VMinutos1 = 60 - Mid(MskTurIni.Text, 4, 2)
                    'MINUTOS DESPUES DE LA MADRUGADA
                    VMinutos2 = Mid(MskTurFin.Text, 4, 2)
                    'SUMA LOS DOS TIPOS DE MINUTOS
                    VMinutos = (VMinutos1 + VMinutos2)
                    
                    If VMinutos > 60 Then
                        VHoras = VHoras + 1
                        VMinutos = VMinutos - 60
                    Else
                        VMinutos = VMinutos / 60
                    End If
                        
                    If (VMinutos1 + VMinutos2) > 60 Then
                        TxtTexto.Item(11).Text = VHoras & "." & VMinutos
                    Else
                        TxtTexto.Item(11).Text = VHoras + VMinutos
                    End If
                Else
                    'RESTA LAS HORAS
                    VHoras = (Mid(MskTurFin.Text, 1, 2) - Mid(MskTurIni.Text, 1, 2))
                    'RESTA LOS MINUTOS
                    VMinutos = (Mid(MskTurFin.Text, 4, 2) - Mid(MskTurIni.Text, 4, 2))
                    
                    If VMinutos > 0 Then
                        TxtTexto.Item(11).Text = VHoras & "." & VMinutos
                    Else
                        'SUMA TODOS LOS MINUTOS
                        TxtTexto.Item(11).Text = VHoras
                    End If
                    
                    
                End If
        End If
End Sub

Private Sub MskTurIni_GotFocus()
        MskTurIni.SelStart = 0
        MskTurIni.SelLength = Len(MskTurIni.Text)
End Sub

Private Sub MskTurIni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub txtbuscli_Change()

        'CODIGO
        If OptProCod.Value = True Then
                    'CUALQUIER PALABRA
                        If BLinea = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Linea,  Descrip from Lineas where Linea Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                        Criteria = "Select Linea,  Descrip from Lineas where UPPER(Linea) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BFicha = True Or BFicha2 = True Or BFicha3 = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica where Esp_Tec Like '%" & TxtBuscli.Text & "%' And Activa = -1"
                                Else 'ORACLE
                                        Criteria = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscli.Text) & "%' And Activa = -1"
                                End If
                        ElseIf BParo = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select CodigoParo, DescripcionParo, Tipo from Paros where CodigoParo Like '%" & TxtBuscli.Text & "%' Order By DescripcionParo"
                                Else 'ORACLE
                                        Criteria = "Select CodigoParo, DescripcionParo, Tipo from Paros where UPPER(CodigoParo) Like '%" & UCase(TxtBuscli.Text) & "%' Order By DescripcionParo"
                                End If
                        ElseIf BMateriaPrima = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Esp_Tec, Descrip from FichaTecnica where Esp_Tec Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                        Criteria = "Select Esp_Tec, Descrip from FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BUnidadMedida = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from UnidadesMedida where Codigo Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                    Criteria = "Select Codigo, Descripcion from UnidadesMedida where UPPER(Codigo) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BEmpleado = True Or BEmpleado2 = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from Empleados where Codigo Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from Empleados where UPPER(Codigo) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        
                        ElseIf BPasada = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from Pasadas where Codigo Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from Pasadas where UPPER(Codigo) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BEquipos = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from EmpleadosGrupos where Codigo Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from EmpleadosGrupos where UPPER(Codigo) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BGrupo = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno where Linea Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno where UPPER(Linea) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        End If
                    
        'DESCRIPCION
        ElseIf OptProDes.Value = True Then
                    'CUALQUIER PALABRA
                        If BLinea = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Linea,  Descrip from Lineas where Descrip Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                        Criteria = "Select Linea,  Descrip from Lineas where UPPER(Descrip) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BFicha = True Or BFicha2 = True Or BFicha3 = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica where Descrip Like '%" & TxtBuscli.Text & "%' And Activa = -1"
                                Else 'ORACLE
                                        Criteria = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBuscli.Text) & "%' And Activa = -1"
                                End If
                        ElseIf BParo = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select CodigoParo, DescripcionParo, Tipo from Paros where DescripcionParo Like '%" & TxtBuscli.Text & "%' Order By DescripcionParo"
                                Else 'ORACLE
                                        Criteria = "Select CodigoParo, DescripcionParo, Tipo from Paros where UPPER(DescripcionParo) Like '%" & UCase(TxtBuscli.Text) & "%' Order By DescripcionParo"
                                End If
                        ElseIf BMateriaPrima = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Esp_Tec, Descrip from FichaTecnica where Descrip Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                        Criteria = "Select Esp_Tec, Descrip from FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BUnidadMedida = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from UnidadesMedida where Descripcion Like '%" & TxtBuscli.Text & "%'"
                                Else 'ORACLE
                                        Criteria = "Select Codigo, Descripcion from UnidadesMedida where UPPER(Descripcion) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BEmpleado = True Or BEmpleado2 = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from Empleados where Descripcion Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from Empleados where UPPER(Descripcion) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BPasada = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from Pasadas where Descripcion Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from Pasadas where UPPER(Descripcion) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BEquipos = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Codigo, Descripcion from EmpleadosGrupos where Descripcion Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Codigo, Descripcion from EmpleadosGrupos where UPPER(Descripcion) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        ElseIf BGrupo = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Criteria = "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno where Linea Like '%" & TxtBuscli.Text & "%'"
                                Else
                                        Criteria = "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno where UPPER(Linea) Like '%" & UCase(TxtBuscli.Text) & "%'"
                                End If
                        End If
                        
                        
        End If
                
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "" & Criteria)
            'LLENA EL GRID
            Set DbGridBusqueda.DataSource = RBusqueda
            If BGrupo = True Then
            Else
                DbGridBusqueda.Columns(1).Width = "4000"
            End If

End Sub

Private Sub TxtBuscli_GotFocus()
    TxtBuscli.SelStart = 0
    TxtBuscli.SelLength = Len(TxtBuscli.Text)
End Sub

Private Sub TxtBuscli_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub txtbuscli_LostFocus()
        TxtBuscli = UCase(TxtBuscli)
End Sub



Private Sub TxtDoc_GotFocus()
        TxtDoc.SelStart = 0
        TxtDoc.SelLength = Len(TxtDoc.Text)
        
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLinCon_GotFocus()
        TxtLinCon.SelStart = 0
        TxtLinCon.SelLength = Len(TxtLinCon.Text)
End Sub

Private Sub TxtLinCon_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
On Error Resume Next

                                    'ORDEN EN CAPTURA DE PAROS
                                    If Index = 6 Then
                                        Set RBuscaFichaOrden = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaFichaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(6).Text & "'")
                                            If RBuscaFichaOrden.RecordCount > 0 Then
                                                TxtTexto.Item(7).Text = RBuscaFichaOrden!FichaTecnica
                                            Else
                                                TxtTexto.Item(7).Text = ""
                                            End If
                                    End If
                          
                                    'FICHA TECNICA DE PAROS
                                    If Index = 7 Then
                                                Set RBuscaFicha = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(7).Text & "'")
                                                    If RBuscaFicha.RecordCount > 0 Then
                                                            LblFicha.Caption = RBuscaFicha(0)
                                                    Else
                                                            LblFicha.Caption = ""
                                                    End If
                                    End If
                                    
                                    'BUSCA EMPLEADO
                                    If Index = 8 Then
                                                Set RBuscaEmpleado = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtTexto.Item(8).Text & "'")
                                                    If RBuscaEmpleado.RecordCount > 0 Then
                                                        LblEmpleado.Caption = RBuscaEmpleado!Descripcion
                                                    Else
                                                        LblEmpleado.Caption = ""
                                                    End If
                                                    
                                    End If
        
                                    'ORDEN EN CONSUMOS DE MATERIAS PRIMAS
                                    If Index = 4 Then
                                        Set RBuscaFichaOrden2 = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaFichaOrden2, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(4).Text & "'")
                                            If RBuscaFichaOrden2.RecordCount > 0 Then
                                                TxtTexto.Item(3).Text = RBuscaFichaOrden2!FichaTecnica
                                            Else
                                                TxtTexto.Item(3).Text = ""
                                            End If
                                    End If
        
                                    'FICHA TECNICA DE CONSUMOS
                                    If Index = 3 Then
                                                Set RBuscaFicha2 = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaFicha2, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(3).Text & "'")
                                                    If RBuscaFicha2.RecordCount > 0 Then
                                                            LblFicha2.Caption = RBuscaFicha2(0)
                                                    Else
                                                            LblFicha2.Caption = ""
                                                    End If
                                    End If
        
                                    'DETALLE DE ORDEN
                                    If Index = 9 Then
                                        Set RBuscaFichaOrden3 = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaFichaOrden3, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(9).Text & "'")
                                            If RBuscaFichaOrden3.RecordCount > 0 Then
                                                TxtTexto.Item(12).Text = RBuscaFichaOrden3!FichaTecnica
                                            Else
                                                TxtTexto.Item(12).Text = ""
                                            End If
                                            
                                            'BUSCA LA INFORMACION DEL DETALLE DE LA ORDEN
                                            Set RInformacion = New ADODB.Recordset
                                            Call Abrir_Recordset(RInformacion, "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo, DO.Desperdicio From DetalleOrdenProduccion DO, Lineas L, Pasadas P Where DO.Documento = '" & TxtTexto.Item(9).Text & "' And DO.Linea = L.Linea And DO.Pasada = P.Codigo")
                                            'LLENA EL GRID
                                            Set DBGridInformacion.DataSource = RInformacion
                                            
                                                                                        
                                            DBGridInformacion.Columns(0).Width = "2000"
                                            DBGridInformacion.Columns(1).Width = "1500"
                                            DBGridInformacion.Columns(2).Width = "1200"
                                            DBGridInformacion.Columns(3).Width = "1200"
                                            DBGridInformacion.Columns(4).Width = "1200"
                                            DBGridInformacion.Columns(5).Width = "1200"
                                            DBGridInformacion.Columns(6).Width = "1200"
                                            DBGridInformacion.Columns(3).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(4).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(5).NumberFormat = "#,###,##0"
                                            DBGridInformacion.Columns(6).NumberFormat = "#,###,##0"
                                    End If
                                                                
                                    'FICHA TECNICA DE PRODUCCION
                                    If Index = 12 Then
                                                Set RBuscaFicha = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(12).Text & "'")
                                                If RBuscaFicha.RecordCount > 0 Then
                                                        LblFicha3.Caption = RBuscaFicha(0)
                                                Else
                                                        LblFicha3.Caption = ""
                                                End If
                                    End If
                                    
                                    'PASADAS
                                    If Index = 13 Then
                                        Set RBuscaPasada = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaPasada, "Select Descripcion From Pasadas Where Codigo = '" & TxtTexto.Item(13).Text & "'")
                                            If RBuscaPasada.RecordCount > 0 Then
                                                LblPasada.Caption = RBuscaPasada!Descripcion
                                            Else
                                                LblPasada.Caption = ""
                                            End If
                                    End If

                                    'EQUIPOS
                                    If Index = 14 Then
                                        Set RBuscaEquipos = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaEquipos, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtTexto.Item(14).Text & "'")
                                            If RBuscaEquipos.RecordCount > 0 Then
                                                LblEquipo.Caption = RBuscaEquipos!Descripcion
                                            Else
                                                LblEquipo.Caption = ""
                                            End If
                                    End If
                                    
                                    'BUSCA EMPLEADO
                                    If Index = 15 Then
                                                Set RBuscaEmpleado = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtTexto.Item(15).Text & "'")
                                                    If RBuscaEmpleado.RecordCount > 0 Then
                                                        LblEmp.Caption = RBuscaEmpleado!Descripcion
                                                    Else
                                                        LblEmp.Caption = ""
                                                    End If
                                                    
                                    End If
             
                                    'LINEAS
                                    If Index = 0 Then
                                                Set RBuscaLinea = New ADODB.Recordset
                                                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
                                                If RBuscaLinea.RecordCount > 0 Then
                                                        LblLinea.Caption = RBuscaLinea(0)
                                                Else
                                                        LblLinea.Caption = ""
                                                End If
                                    End If
        
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
                       
            'INICIALIZA EL RECORDSET
            Set RBusqueda = New ADODB.Recordset
            
            'LINEAS
            If Index = 0 Then
                    BLinea = True
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Orden, Velocidad From Lineas")
            'FICHA TECNICA DE PAROS
            ElseIf Index = 7 Then
                    BLinea = False
                    BFicha = True
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
            'FICHA TECNICA DE CONSUMOS
            ElseIf Index = 3 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = True
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
            'EMPLEADOS
            ElseIf Index = 8 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = True
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados")
            'FICHA TECNICA DE PRODUCCION
            ElseIf Index = 12 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = True
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
            'PASADAS
            ElseIf Index = 13 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = True
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas")
            'EQUIPOS
            ElseIf Index = 14 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = True
                    BEmpleado2 = False
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos")
            'EQUIPOS
            ElseIf Index = 15 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = True
                    BGrupo = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados")
            'GRUPOS
            ElseIf Index = 24 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    BPasada = False
                    BEquipos = False
                    BEmpleado2 = False
                    BGrupo = True
                    Call Abrir_Recordset(RBusqueda, "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno")
            End If
            
            If (Index = 0 Or Index = 7 Or Index = 3 Or Index = 8 Or Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Or Index = 24) Then
                    
                    Set DbGridBusqueda.DataSource = RBusqueda
                    If Index = 24 Then
                        DbGridBusqueda.Columns(1).Width = "1000"
                    Else
                        DbGridBusqueda.Columns(1).Width = "4000"
                    End If
                    FrameBusqueda.Visible = True
                    TxtBuscli.SetFocus
            End If
            
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        TxtTexto.Item(Index).SelStart = 0
        TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
            If KeyAscii = 13 Then
                            SendKeys "{tab}", True
            End If
            
            If KeyAscii = 43 Then
                            'INICIALIZA EL RECORDSET
                                Set RBusqueda = New ADODB.Recordset
                                
                                'LINEAS
                            If Index = 0 Then
                                    BLinea = True
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Orden, Velocidad From Lineas")
                            'FICHA TECNICA DE PAROS
                            ElseIf Index = 7 Then
                                    BLinea = False
                                    BFicha = True
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
                            'FICHA TECNICA DE CONSUMOS
                            ElseIf Index = 3 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = True
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
                            'EMPLEADOS
                            ElseIf Index = 8 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = True
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados")
                            'FICHA TECNICA DE PRODUCCION
                            ElseIf Index = 12 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = True
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
                            'PASADAS
                            ElseIf Index = 13 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = True
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas")
                            'EQUIPOS
                            ElseIf Index = 14 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = True
                                    BEmpleado2 = False
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos")
                            'EQUIPOS
                            ElseIf Index = 15 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = True
                                    BGrupo = False
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados")
                            'GRUPOS
                            ElseIf Index = 24 Then
                                    BLinea = False
                                    BFicha = False
                                    BParo = False
                                    BMateriaPrima = False
                                    BFicha2 = False
                                    BUnidadMedida = False
                                    BEmpleado = False
                                    BFicha3 = False
                                    BPasada = False
                                    BEquipos = False
                                    BEmpleado2 = False
                                    BGrupo = True
                                    Call Abrir_Recordset(RBusqueda, "Select Linea, OperadorEntrega, MecanicoEntrega, InspectorEntrega, SupervisorEntrega, OperadorRecibe, MecanicoRecibe, InspectorRecibe, SupervisorRecibe  from LineasPersonalTurno")
                            End If
                            
                            If (Index = 0 Or Index = 7 Or Index = 3 Or Index = 8 Or Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Or Index = 24) Then
                                    
                                    Set DbGridBusqueda.DataSource = RBusqueda
                                    If Index = 24 Then
                                        DbGridBusqueda.Columns(1).Width = "1000"
                                    Else
                                        DbGridBusqueda.Columns(1).Width = "4000"
                                    End If
                                    FrameBusqueda.Visible = True
                                    TxtBuscli.SetFocus
                            End If
            
                End If
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)

    'GRUPO
    If Index = 24 Then
        If TxtTexto.Item(24).Text <> "" Then
            Set RLineaActiva = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RLineaActiva, "Select * From LineasPersonalTurno Where Linea = '" & TxtTexto.Item(24).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RLineaActiva, "Select * From LineasPersonalTurno Where UPPER(Linea) = '" & UCase(TxtTexto.Item(24).Text) & "'")
            End If
                  If RLineaActiva.RecordCount > 0 Then
                     If IsNull(RLineaActiva!OperadorEntrega) Then
                        TxtTexto.Item(16).Text = ""
                     Else
                        TxtTexto.Item(16).Text = RLineaActiva!OperadorEntrega
                     End If
                     If IsNull(RLineaActiva!MecanicoEntrega) Then
                        TxtTexto.Item(17).Text = ""
                     Else
                        TxtTexto.Item(17).Text = RLineaActiva!MecanicoEntrega
                     End If
                     If IsNull(RLineaActiva!InspectorEntrega) Then
                        TxtTexto.Item(18).Text = ""
                     Else
                        TxtTexto.Item(18).Text = RLineaActiva!InspectorEntrega
                     End If
                     If IsNull(RLineaActiva!SupervisorEntrega) Then
                        TxtTexto.Item(19).Text = ""
                     Else
                        TxtTexto.Item(19).Text = RLineaActiva!SupervisorEntrega
                     End If
                     If IsNull(RLineaActiva!OperadorRecibe) Then
                        TxtTexto.Item(20).Text = ""
                     Else
                        TxtTexto.Item(20).Text = RLineaActiva!OperadorRecibe
                     End If
                     If IsNull(RLineaActiva!MecanicoRecibe) Then
                        TxtTexto.Item(21).Text = ""
                     Else
                        TxtTexto.Item(21).Text = RLineaActiva!MecanicoRecibe
                     End If
                     If IsNull(RLineaActiva!InspectorRecibe) Then
                        TxtTexto.Item(22).Text = ""
                     Else
                        TxtTexto.Item(22).Text = RLineaActiva!InspectorRecibe
                     End If
                     If IsNull(RLineaActiva!SupervisorRecibe) Then
                        TxtTexto.Item(23).Text = ""
                     Else
                        TxtTexto.Item(23).Text = RLineaActiva!SupervisorRecibe
                     End If
                  Else
                     
                     TxtTexto.Item(16).Text = ""
                     TxtTexto.Item(17).Text = ""
                     TxtTexto.Item(18).Text = ""
                     TxtTexto.Item(19).Text = ""
                     TxtTexto.Item(20).Text = ""
                     TxtTexto.Item(21).Text = ""
                     TxtTexto.Item(22).Text = ""
                     TxtTexto.Item(23).Text = ""
                  End If
        End If
    'BUSCA TURNO
    ElseIf Index = 2 Then
        Set RBuscaTurno = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTurno, "Select * From Turnos Where Turno = '" & TxtTexto.Item(2).Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaTurno, "Select * From Turnos Where UPPER(Turno) = '" & UCase(TxtTexto.Item(2).Text) & "'")
        End If
            If RBuscaTurno.RecordCount > 0 Then
                    MskTurIni.Text = RBuscaTurno!Inicio
                    MskTurFin.Text = RBuscaTurno!Termina
            End If
                
    End If
               
End Sub

Private Sub TxtTexto2_Change(Index As Integer)
On Error Resume Next
                                'BUSCA EL CODIGO DE PARO
                                    If Index = 4 Then
                                                Set RBuscaParo = New ADODB.Recordset
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RBuscaParo, "Select P.DescripcionParo, P.Tipo, PG.Descripcion From Paros P, ParosGrupos PG Where P.CodigoParo = '" & TxtTexto2.Item(4).Text & "' And P.Grupo = PG.CodigoGrupo")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RBuscaParo, "Select P.DescripcionParo, P.Tipo, PG.Descripcion From Paros P, ParosGrupos PG Where UPPER(P.CodigoParo) = '" & UCase(TxtTexto2.Item(4).Text) & "' And P.Grupo = PG.CodigoGrupo")
                                                End If
                                                If RBuscaParo.RecordCount > 0 Then
                                                                LblParo.Caption = RBuscaParo(0)
                                                                LblTipo.Caption = RBuscaParo(1)
                                                                LblParGru.Caption = RBuscaParo(2)
                                                                
                                                Else
                                                                LblParo.Caption = ""
                                                                LblTipo.Caption = ""
                                                                LblParGru.Caption = ""
                                                End If
                                    End If

            
                                    'MATERIA PRIMA
                                    If Index = 5 Then
                                                Set RBuscaMateriaPrima = New ADODB.Recordset
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip, UnidadMedida From FichaTecnica Where Esp_Tec = '" & TxtTexto2.Item(5).Text & "'")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip, UnidadMedida From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto2.Item(5).Text) & "'")
                                                End If
                                                
                                                If RBuscaMateriaPrima.RecordCount > 0 Then
                                                                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descrip
                                                                LblUnidadMedida.Caption = RBuscaMateriaPrima!unidadMedida
                                                Else
                                                                LblMateriaPrima.Caption = ""
                                                                LblUnidadMedida.Caption = ""
                                                                
                                                End If
                                    End If
            
                    
                                    'UNIDAD DE MEDIDA
                                    'If Index = 7 Then
                                    '            Set RBuscaUnidadMedida = New ADODB.Recordset
                                    '            If GOrigenDeDatos = "AmaproAccess" Then
                                    '                Call Abrir_Recordset(RBuscaUnidadMedida, "Select Descripcion From UnidadesMedida Where Codigo = '" & TxtTexto2.Item(7).Text & "'")
                                    '            Else 'ORACLE
                                    '                Call Abrir_Recordset(RBuscaUnidadMedida, "Select Descripcion From UnidadesMedida Where Upper(Codigo) = '" & UCase(TxtTexto2.Item(7).Text) & "'")
                                    '            End If
                                    '
                                    '            If RBuscaUnidadMedida.RecordCount > 0 Then
                                    '                            LblUnidadMedida.Caption = RBuscaUnidadMedida!Descripcion
                                    '            Else
                                    '                            LblUnidadMedida.Caption = ""
                                    '
                                    '            End If
                                    
                                   ' End If
End Sub

Private Sub TxtTexto2_DblClick(Index As Integer)
On Error Resume Next
            Set RBusqueda = New ADODB.Recordset
            'PAROS
            If Index = 4 Then
                    BLinea = False
                    BFicha = False
                    BParo = True
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoParo, DescripcionParo, Tipo from Paros Order By DescripcionParo")
                    Set DbGridBusqueda.DataSource = RBusqueda
                    DbGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBuscli.SetFocus
            'MATERIA PRIMA
            ElseIf Index = 5 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = True
                    BFicha2 = False
                    BUnidadMedida = False
                    BEmpleado = False
                    BFicha3 = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
                    Set DbGridBusqueda.DataSource = RBusqueda
                    DbGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBuscli.SetFocus
            'UNIDAD DE MEDIDA
            ElseIf Index = 7 Then
                    BLinea = False
                    BFicha = False
                    BParo = False
                    BMateriaPrima = False
                    BFicha2 = False
                    BUnidadMedida = True
                    BEmpleado = False
                    BFicha3 = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from UnidadesMedida")
                    Set DbGridBusqueda.DataSource = RBusqueda
                    DbGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBuscli.SetFocus
            End If
            
End Sub

Private Sub TxtTexto2_GotFocus(Index As Integer)
On Error Resume Next
    If Index > 0 Then
        TxtTexto2.Item(Index).SelStart = 0
        TxtTexto2.Item(Index).SelLength = Len(TxtTexto2.Item(Index).Text)
    End If
End Sub

Private Sub TxtTexto2_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
        If Index > 0 Then
            If KeyAscii = 13 Then
                        SendKeys "{TAB}"
                        
                        If Err > 0 Then
                            MsgBox Err.Number & Err.Description
                        End If
            End If
        End If
            
        
        If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    'PAROS
                    If Index = 4 Then
                            BLinea = False
                            BFicha = False
                            BParo = True
                            BMateriaPrima = False
                            BFicha2 = False
                            BUnidadMedida = False
                            BEmpleado = False
                            BFicha3 = False
                            Call Abrir_Recordset(RBusqueda, "Select CodigoParo, DescripcionParo, Tipo from Paros Order By DescripcionParo")
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBuscli.SetFocus
                    'MATERIA PRIMA
                    ElseIf Index = 5 Then
                            BLinea = False
                            BFicha = False
                            BParo = False
                            BMateriaPrima = True
                            BFicha2 = False
                            BUnidadMedida = False
                            BEmpleado = False
                            BFicha3 = False
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBuscli.SetFocus
                    'UNIDAD DE MEDIDA
                    ElseIf Index = 7 Then
                            BLinea = False
                            BFicha = False
                            BParo = False
                            BMateriaPrima = False
                            BFicha2 = False
                            BUnidadMedida = True
                            BEmpleado = False
                            BFicha3 = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from UnidadesMedida")
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBuscli.SetFocus
                            
                    End If
        End If

End Sub


Public Sub BotonesVisiblesEncabezado()
    If BanderaBotonesVisiblesEncabezado = True Then
        CmdAgregar.Visible = True
        CmdEditar.Visible = True
        CmdBorrar.Visible = True
        CmdGrabar.Visible = True
        CmdCancelar.Visible = True
        CmdBuscar.Visible = True
        CmdSalida.Visible = True
        CmdImprimir.Visible = True
    Else
        CmdAgregar.Visible = False
        CmdEditar.Visible = False
        CmdBorrar.Visible = False
        CmdGrabar.Visible = False
        CmdCancelar.Visible = False
        CmdBuscar.Visible = False
        CmdSalida.Visible = False
        CmdImprimir.Visible = False
    End If

End Sub


Public Sub Botones3()
    If Bandera3 = True Then
         FrameDetalleMateriaPrima.Enabled = True
         CmdBotones3.Item(0).Enabled = False
         'CmdBotones3.Item(1).Enabled = False
         CmdBotones3.Item(2).Enabled = True
         CmdBotones3.Item(3).Enabled = True
         CmdBotones3.Item(4).Enabled = False
         CmdBotones3.Item(5).Enabled = False
    Else
         FrameDetalleMateriaPrima.Enabled = False
         CmdBotones3.Item(0).Enabled = True
         'CmdBotones3.Item(1).Enabled = True
         CmdBotones3.Item(2).Enabled = False
         CmdBotones3.Item(3).Enabled = False
         CmdBotones3.Item(4).Enabled = True
         CmdBotones3.Item(5).Enabled = True
    End If

End Sub

Public Sub BotonesVisibles2()
    If BanderaBotonesVisibles2 = True Then
        CmdBotones3.Item(0).Visible = True
        'CmdBotones3.Item(1).Visible = True
        CmdBotones3.Item(2).Visible = True
        CmdBotones3.Item(3).Visible = True
        CmdBotones3.Item(4).Visible = True
        CmdBotones3.Item(5).Visible = True
    Else
        CmdBotones3.Item(0).Visible = False
        'CmdBotones3.Item(1).Visible = False
        CmdBotones3.Item(2).Visible = False
        CmdBotones3.Item(3).Visible = False
        CmdBotones3.Item(4).Visible = False
        CmdBotones3.Item(5).Visible = False
    End If

End Sub

Public Sub Botones4()
    If Bandera4 = True Then
         FrameDetalleProduccion.Enabled = True
         CmdBotones4.Item(0).Enabled = False
         'CmdBotones4.Item(1).Enabled = False
         CmdBotones4.Item(2).Enabled = True
         CmdBotones4.Item(3).Enabled = True
         CmdBotones4.Item(4).Enabled = False
         CmdBotones4.Item(5).Enabled = False
    Else
         FrameDetalleProduccion.Enabled = False
         CmdBotones4.Item(0).Enabled = True
         'CmdBotones4.Item(1).Enabled = True
         CmdBotones4.Item(2).Enabled = False
         CmdBotones4.Item(3).Enabled = False
         CmdBotones4.Item(4).Enabled = True
         CmdBotones4.Item(5).Enabled = True
    End If

End Sub

Public Sub BotonesVisibles3()
    If BanderaBotonesVisibles3 = True Then
        CmdBotones4.Item(0).Visible = True
        'CmdBotones4.Item(1).Visible = True
        CmdBotones4.Item(2).Visible = True
        CmdBotones4.Item(3).Visible = True
        CmdBotones4.Item(4).Visible = True
        CmdBotones4.Item(5).Visible = True
    Else
        CmdBotones4.Item(0).Visible = False
        'CmdBotones4.Item(1).Visible = False
        CmdBotones4.Item(2).Visible = False
        CmdBotones4.Item(3).Visible = False
        CmdBotones4.Item(4).Visible = False
        CmdBotones4.Item(5).Visible = False
    End If

End Sub

Private Sub TxtTexto2_LostFocus(Index As Integer)
        If Index = 1 Then
        
                    'SI ESTA EN BLANCO SE BUSCA LA FECHA Y LINEA SI NO DEJA LA QUE ESTAS
                    If IsNumeric(TxtTexto2.Item(1).Text) Then
                        If MskFecCon.Text = "" Then
                                Set RBuscaTarima = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTexto2.Item(1).Text & " And FichaTecnica = '" & TxtTexto2.Item(5).Text & "' And Linea = '" & TxtLinCon.Text & "'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTexto2.Item(1).Text & " And UPPER(FichaTecnica) = '" & UCase(TxtTexto2.Item(5).Text) & "' And UPPER(Linea) = '" & UCase(TxtLinCon.Text) & "'")
                                    End If
                                
                                If RBuscaTarima.RecordCount > 0 Then
                                        MskFecCon.Text = RBuscaTarima!FechaProduccion
                                Else
                                    MsgBox "Ficha Tecnica Con Este Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                                    Exit Sub
                                End If
                        End If
                    End If
        
        
        
                        'REVISA SI EXISTE EL BULTO
                        Set RBuscaSaldo = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaSaldo, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = #" & Format(MskFecCon.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLinCon.Text & "' And FichaTecnica = '" & TxtTexto2.Item(5).Text & "' And Tarima = " & TxtTexto2.Item(1).Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaSaldo, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = TO_DATE('" & MskFecCon.Text & "','dd/mm/yyyy')" & " And Linea = '" & TxtLinCon.Text & "' And FichaTecnica = '" & TxtTexto2.Item(5).Text & "' And Tarima = " & TxtTexto2.Item(1).Text)
                            End If
                
                            If RBuscaSaldo.RecordCount > 0 Then
                                TxtSaldo.Text = RBuscaSaldo!Saldo
                            Else
                                TxtSaldo.Text = "Bulto No Existe"
                            End If
        End If
End Sub

Public Sub CalculaEficiencia()

            
                        
                'BUSCA LOS DATOS DEL DOCUMENTO
                Set RCapturaParos = New ADODB.Recordset
                Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Documento = " & VDocumento)
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                         Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal From EncabezadoCapturaParos Where Documento = " & VDocumento)
                         
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                             Else
                                VTiempoProgramadoD = 0
                                VVelocidadTeoricaDia = 0
                                VVelocidadRealDia = 0
                             End If
                        
                                               
                                                                                                
                 '********  PAROS QUE NO AFECTAN 'N' **************************************************************
                        
                        'BUSCAR PAROS QUE NO AFECTAN DEL TURNO DE DIA
                        Set RBuscaParosNoAfectanD = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Tipo = 'N'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And UPPER(P.Tipo) = 'N'")
                        End If
                                            
                            If RBuscaParosNoAfectanD.RecordCount > 0 Then
                                If IsNull(RBuscaParosNoAfectanD(0)) Then
                                    VParosND = 0
                                Else
                                    VParosND = RBuscaParosNoAfectanD(0) / 60
                                End If
                            Else
                                VParosND = 0
                            End If
                '********  PAROS QUE SI AFECTAN 'S' **************************************************************
                                            
                                            
                        'BUSCAR PAROS QUE AFECTAN DEL TURNO DE DIA
                        Set RBuscaParosSiAfectanD = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Tipo = 'S'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And UPPER(P.Tipo) = 'S'")
                        End If
                            If RBuscaParosSiAfectanD.RecordCount > 0 Then
                                If IsNull(RBuscaParosSiAfectanD(0)) Then
                                    VParosSD = 0
                                Else
                                    VParosSD = RBuscaParosSiAfectanD(0) / 60
                                End If
                            Else
                                VParosSD = 0
                            End If
                                            
                           
                '********  PRODUCCION **************************************************************
                                            
                                            
                        'BUSCAR PRODUCCION DIA
                         Set RBuscaProduccionD = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And P.Tipo = 'P'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, Paros P Where DP.Documento = " & VDocumento & " And DP.Paro = P.CodigoParo And UPPER(P.Tipo) = 'P'")
                        End If
                                            
                            If RBuscaProduccionD.RecordCount > 0 Then
                                If IsNull(RBuscaProduccionD(0)) Then
                                    VProduccionD = 0
                                Else
                                    VProduccionD = RBuscaProduccionD(0) / 60
                                End If
                            Else
                                VProduccionD = 0
                            End If
                                            
    '***************************************************************************************************************
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        
    'DIA _______________________________________________________________________________________________________
                        
             'PRODUCTO CONFORME
                        'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
                        Set RProduccion = New ADODB.Recordset
                        Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Documento = " & VDocumento)
                                                             
                            If RProduccion.RecordCount > 0 Then
                                If IsNull(RProduccion(0)) Then
                                    VPCD = 0
                                Else
                                    VPCD = RProduccion(0)
                                End If
                            Else
                                VPCD = 0
                            End If
                                                             
            'PRODUCTO NO CONFORME
                        'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
                        Set RProduccion = New ADODB.Recordset
                        Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Documento = " & VDocumento)
                                                                
                            If RProduccion.RecordCount > 0 Then
                                If IsNull(RProduccion(0)) Then
                                    VPNCD = 0
                                Else
                                    VPNCD = RProduccion(0)
                                End If
                            Else
                                    VPNCD = 0
                            End If
                                                             
            'DESPERDICIO
                        Set RProduccion = New ADODB.Recordset
                        Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Documento = " & VDocumento)
                                                                                                                                            
                            If RProduccion.RecordCount > 0 Then
                                If IsNull(RProduccion(0)) Then
                                    VPDD = 0
                                Else
                                    VPDD = RProduccion(0)
                                End If
                            Else
                                VPDD = 0
                            End If
                        
                        
                        VTotalProduccionD = VPCD + VPNCD
                        
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
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
                                        
                'FACTOR 2______________________________________________________________________
                            If VPNCD > VTotalProduccionD Then
                                VFactor2D = 0
                            Else
                                    VFactor2D = VTotalProduccionD - VPNCD
                                    If VTotalProduccionD = 0 Then
                                    Else
                                       VFactor2D = VFactor2D / VTotalProduccionD
                                    End If
                            End If
                                                
                'FACTOR 3______________________________________________________________________
                            If VPDD > VTotalProduccionD Then
                                VFactor3D = 0
                            Else
                                    VFactor3D = VTotalProduccionD - VPDD
                                    If VTotalProduccionD = 0 Then
                                    Else
                                       VFactor3D = VFactor3D / VTotalProduccionD
                                    End If
                            End If
                                        
                'FACTOR 4______________________________________________________________________
                            VFactor4D = VTiempoProgramadoD - VParosND
                            VFactor4D = VFactor4D - VParosSD
                            If (VTiempoProgramadoD - VParosND) = 0 Then
                            Else
                                VFactor4D = (VFactor4D / (VTiempoProgramadoD - VParosND))
                            End If
                                                        
                                        
                'FACTOR 5______________________________________________________________________
                            If VVelocidadTeoricaDia = 0 Then
                                VFactor5D = 0
                            Else
                                VFactor5D = VVelocidadRealDia / VVelocidadTeoricaDia
                            End If
                                                
                
                'EFICIENCIA REAL DEL TURNO DE DIA
                             VEficienciaRealD = VFactor1D * VFactor2D * VFactor3D * VFactor4D * VFactor5D * 100
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                       ' If BMenorHorasD = True Then
                                       '     VTiempoProgramadoD = VTiempoProgramadoD / 100
                                       ' Else
                                            VTiempoProgramadoD = VTiempoProgramadoD / 60
                                       ' End If
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
                                    
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        'Exit Sub
                                    End If

                                    'ACTUALIZA Y GRABA LA EFICIENCIA DEL REPORTE
                                    Conexion.Execute ("update EncabezadoCapturaParos Set Eficiencia = " & Format(VEficienciaRealD, "#,###,##0.0000") & ", EficienciaReporte = " & Format(VFactor1D, "#,###,##0.0000") & " where documento = " & VDocumento)
                                    
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                    End If
                                    
                            
                    End If
                


End Sub

Public Sub Llena_CamposEncabezado()
On Error Resume Next
    If REncabezadoParos.RecordCount > 0 Then
        'DOCUMENTO
        TxtDoc.Text = REncabezadoParos!Documento
        'FECHA
        MskFec.Text = REncabezadoParos!fecha
        'TURNO
            If IsNull(REncabezadoParos!Turno) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = REncabezadoParos!Turno
            End If
        'INICIO
            If IsNull(REncabezadoParos!Inicio) Then
                MskTurIni.Text = ""
            Else
                MskTurIni.Text = REncabezadoParos!Inicio
            End If
        'TRMINA
            If IsNull(REncabezadoParos!Termina) Then
                MskTurFin.Text = ""
            Else
                MskTurFin.Text = REncabezadoParos!Termina
            End If
        'HORAS PROGRAMADAS
            If IsNull(REncabezadoParos!HorasProgramadas) Then
                TxtTexto.Item(11).Text = ""
            Else
                TxtTexto.Item(11).Text = REncabezadoParos!HorasProgramadas
            End If
        'USUARIO
            If IsNull(REncabezadoParos!Usuario) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = REncabezadoParos!Usuario
            End If
            'PRODUCTO CONFORME
            MskProducto.Item(0).Text = REncabezadoParos!ProductoConforme
            'PNC
            MskProducto.Item(1).Text = REncabezadoParos!ProductoNoConforme
            'PRODUCTO EN PROCESO
            MskProducto.Item(2).Text = REncabezadoParos!ProductoEnProceso
            'DESPERDICIO
            MskProducto.Item(3).Text = REncabezadoParos!Desperdicio
            'VELOCIDAD REAL
            TxtTexto.Item(10).Text = REncabezadoParos!VelocidadReal
            'VELOCIDAD TEORICA
            TxtTexto.Item(1).Text = REncabezadoParos!VelocidadTeorica
            'LINEA
            If IsNull(REncabezadoParos!Linea) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = REncabezadoParos!Linea
            End If
            'GRUPO
            If IsNull(REncabezadoParos!Grupo) Then
                TxtTexto.Item(14).Text = ""
            Else
                TxtTexto.Item(14).Text = REncabezadoParos!Grupo
            End If
            'EFICIENCIA REPORTE
            
            TxtEfiRep.Text = Format(REncabezadoParos!EficienciaReporte, "#,###,##0.00000")
            'EFICIENCIA
            TxtEficiencia.Text = Format(REncabezadoParos!Eficiencia, "#,###,##0.00")
            
            'OPERADOR ENTREGA
            If IsNull(REncabezadoParos!OperadorEntrega) Then
                TxtTexto.Item(16).Text = ""
            Else
                TxtTexto.Item(16).Text = REncabezadoParos!OperadorEntrega
            End If
            'MECANICO ENTREGA
            If IsNull(REncabezadoParos!MecanicoEntrega) Then
                TxtTexto.Item(17).Text = ""
            Else
                TxtTexto.Item(17).Text = REncabezadoParos!MecanicoEntrega
            End If
            'INSPECTOR ENTREGA
            If IsNull(REncabezadoParos!InspectorEntrega) Then
                TxtTexto.Item(18).Text = ""
            Else
                TxtTexto.Item(18).Text = REncabezadoParos!InspectorEntrega
            End If
            'SUPERVISOR ENTREGA
            If IsNull(REncabezadoParos!SupervisorEntrega) Then
                TxtTexto.Item(19).Text = ""
            Else
                TxtTexto.Item(19).Text = REncabezadoParos!SupervisorEntrega
            End If
            
            'OPERADOR RECIBE
            If IsNull(REncabezadoParos!OperadorRecibe) Then
                TxtTexto.Item(20).Text = ""
            Else
                TxtTexto.Item(20).Text = REncabezadoParos!OperadorRecibe
            End If
            'MECANICO RECIBE
            If IsNull(REncabezadoParos!MecanicoRecibe) Then
                TxtTexto.Item(21).Text = ""
            Else
                TxtTexto.Item(21).Text = REncabezadoParos!MecanicoRecibe
            End If
            'INSPECTOR RECIBE
            If IsNull(REncabezadoParos!InspectorRecibe) Then
                TxtTexto.Item(22).Text = ""
            Else
                TxtTexto.Item(22).Text = REncabezadoParos!InspectorRecibe
            End If
            'SUPERVISOR RECIBE
            If IsNull(REncabezadoParos!SupervisorRecibe) Then
                TxtTexto.Item(23).Text = ""
            Else
                TxtTexto.Item(23).Text = REncabezadoParos!SupervisorRecibe
            End If
    Else
            Limpia_CamposEncabezado
    End If
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_CamposEncabezado()
            'DOCUMENTO
            TxtDoc.Text = ""
            'FECHA
            MskFec.Text = Date
            'TURNO
            TxtTexto.Item(2).Text = ""
            MskTurIni.Text = "__:__"
            MskTurFin.Text = "__:__"
            'HORAS PROGRAMADAS
            TxtTexto.Item(11).Text = 0
            'USUARIO
            TxtTexto.Item(5).Text = ""
            'PRODUCTO CONFORME
            MskProducto.Item(0).Text = 0
            'PNC
            MskProducto.Item(1).Text = 0
            'PRODUCTO EN PROCESO
            MskProducto.Item(2).Text = 0
            'DESPERDICIO
            MskProducto.Item(3).Text = 0
            'VELOCIDAD REAL
            TxtTexto.Item(10).Text = 0
            'VELOCIDAD TEORICA
            TxtTexto.Item(1).Text = 0
            'LINEA
            TxtTexto.Item(0).Text = ""
            'GRUPO
            TxtTexto.Item(14).Text = ""
            'EFICIENCIA REPORTE
            TxtEfiRep.Text = 0
            'EFICIENCIA
            TxtEficiencia.Text = 0
            'OPERADOR ENTREGA
            TxtTexto.Item(16).Text = ""
            'MECANICO ENTREGA
            TxtTexto.Item(17).Text = ""
            'INSPECTOR ENTREGA
            TxtTexto.Item(18).Text = ""
            'SUPERVISOR ENTREGA
            TxtTexto.Item(19).Text = ""
            'OPERADOR RECIBE
            TxtTexto.Item(20).Text = ""
            'MECANICO RECIBE
            TxtTexto.Item(21).Text = ""
            'INSPECTOR RECIBE
            TxtTexto.Item(22).Text = ""
            'SUPERVISOR RECIBE
            TxtTexto.Item(23).Text = ""
        
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
    If RDetalleParos.RecordCount > 0 Then
        'DOCUMENTO
            If IsNull(RDetalleParos!Documento) Then
                TxtTexto2.Item(0).Text = ""
            Else
                TxtTexto2.Item(0).Text = RDetalleParos!Documento
            End If
        'ORDEN
            If IsNull(RDetalleParos!Orden) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = RDetalleParos!Orden
            End If
        'EMPLEADO
            If IsNull(RDetalleParos!Empleado) Then
                TxtTexto.Item(8).Text = ""
            Else
                TxtTexto.Item(8).Text = RDetalleParos!Empleado
            End If
        'INICIO
            If IsNull(RDetalleParos!Inicio) Then
                MskParIni.Text = ""
            Else
                MskParIni.Text = RDetalleParos!Inicio
            End If
        'FINAL
            If IsNull(RDetalleParos!Final) Then
                MskParFin.Text = ""
            Else
                MskParFin.Text = RDetalleParos!Final
            End If
            'MINUTOS
            If IsNull(RDetalleParos!Minutos) Then
                TxtTexto2.Item(3).Text = ""
            Else
                TxtTexto2.Item(3).Text = RDetalleParos!Minutos
            End If
            'PARO
            If IsNull(RDetalleParos!Paro) Then
                TxtTexto2.Item(4).Text = ""
            Else
                TxtTexto2.Item(4).Text = RDetalleParos!Paro
            End If
    Else
            Limpia_CamposDetalle
    End If
            
        If Err <> 0 Then
            
        End If
End Sub

Public Sub Limpia_CamposDetalle()
        'DOCUMENTO
                TxtTexto2.Item(0).Text = ""
        'ORDEN
                TxtTexto.Item(6).Text = ""
        'EMPLEADO
                TxtTexto.Item(8).Text = ""
        'INICIO
                MskParIni.Text = "00:00"
        'FINAL
                MskParFin.Text = "00:00"
        'MINUTOS
                TxtTexto2.Item(3).Text = ""
        'PARO
                TxtTexto2.Item(4).Text = ""
        
End Sub

Public Sub Llena_CamposConsumos()
On Error Resume Next
    If RDetalleConsumos.RecordCount > 0 Then
        'DOCUMENTO
            If IsNull(RDetalleConsumos!Documento) Then
                TxtTexto2.Item(2).Text = ""
            Else
                TxtTexto2.Item(2).Text = RDetalleConsumos!Documento
            End If
        'ORDEN
            If IsNull(RDetalleConsumos!Orden) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RDetalleConsumos!Orden
            End If
        'FECHA
            If IsNull(RDetalleConsumos!fecha) Then
                MskFecCon.Text = ""
            Else
                MskFecCon.Text = RDetalleConsumos!fecha
            End If
        'LINEA
            If IsNull(RDetalleConsumos!Linea) Then
                TxtLinCon.Text = ""
            Else
                TxtLinCon.Text = RDetalleConsumos!Linea
                
            End If
        'FICHA TECNICA
            If IsNull(RDetalleConsumos!FichaTecnica) Then
                TxtTexto2.Item(5).Text = ""
            Else
                TxtTexto2.Item(5).Text = RDetalleConsumos!FichaTecnica
            End If
        'FICHA TECNICA
            If IsNull(RDetalleConsumos!Tarima) Then
                TxtTexto2.Item(1).Text = ""
            Else
                TxtTexto2.Item(1).Text = RDetalleConsumos!Tarima
            End If
            
        'DESPERDICIO
            If IsNull(RDetalleConsumos!Desperdicio) Then
                TxtTexto2.Item(8).Text = ""
            Else
                TxtTexto2.Item(8).Text = RDetalleConsumos!Desperdicio
            End If
        'Cantidad
            If IsNull(RDetalleConsumos!Cantidad) Then
                TxtTexto2.Item(6).Text = ""
            Else
                TxtTexto2.Item(6).Text = RDetalleConsumos!Cantidad
            End If
            'CONTADOR
            If IsNull(RDetalleConsumos!Contador) Then
                TxtConLin.Text = ""
            Else
                TxtConLin.Text = RDetalleConsumos!Contador
            End If
    Else
            Limpia_CamposConsumos
    End If
            
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_CamposConsumos()
        'DOCUMENTO
                TxtTexto2.Item(2).Text = ""
        'ORDEN
                TxtTexto.Item(4).Text = ""
        'FECHA
                MskFecCon.Text = ""
        'LINEA
                
        'FICHA TECNICA
                TxtTexto2.Item(5).Text = ""
        'BULTO O TARIMA
                TxtTexto2.Item(1).Text = ""
        'DESPERDICIO
                TxtTexto2.Item(8).Text = ""
        'Cantidad
                TxtTexto2.Item(6).Text = ""
            
End Sub

Public Sub Llena_CamposProduccion()
On Error Resume Next
    If RDetalleProduccion.RecordCount > 0 Then
        'DOCUMENTO
            If IsNull(RDetalleProduccion!Documento) Then
                TxtTexto2.Item(11).Text = ""
            Else
                TxtTexto2.Item(11).Text = RDetalleProduccion!Documento
            End If
        'ORDEN
            If IsNull(RDetalleProduccion!Orden) Then
                TxtTexto.Item(9).Text = ""
            Else
                TxtTexto.Item(9).Text = RDetalleProduccion!Orden
            End If
        'PASADA
            If IsNull(RDetalleProduccion!Pasada) Then
                TxtTexto.Item(13).Text = ""
            Else
                TxtTexto.Item(13).Text = RDetalleProduccion!Pasada
            End If
        'PC
            If IsNull(RDetalleProduccion!ProductoConforme) Then
                MskProducto.Item(4).Text = 0
            Else
                MskProducto.Item(4).Text = RDetalleProduccion!ProductoConforme
            End If
        'LIBERADO
            If IsNull(RDetalleProduccion!ProductoLiberado) Then
                MskProducto.Item(7).Text = 0
            Else
                MskProducto.Item(7).Text = RDetalleProduccion!ProductoLiberado
            End If
        'PNC
            If IsNull(RDetalleProduccion!ProductoNoConforme) Then
                MskProducto.Item(5).Text = 0
            Else
                MskProducto.Item(5).Text = RDetalleProduccion!ProductoNoConforme
            End If
        'DESPERDICIO
            If IsNull(RDetalleProduccion!Desperdicio) Then
                MskProducto.Item(6).Text = 0
            Else
                MskProducto.Item(6).Text = RDetalleProduccion!Desperdicio
            End If
    Else
            Limpia_CamposProduccion
            
    End If 'FINALIZA RECORDCOUNT
        
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_CamposProduccion()
        'DOCUMENTO
                TxtTexto2.Item(11).Text = ""
        'ORDEN
                TxtTexto.Item(9).Text = ""
        'PASADA
                TxtTexto.Item(13).Text = ""
        'PC
                MskProducto.Item(4).Text = 0
        'PROCESO
                MskProducto.Item(7).Text = 0
        'PNC
                MskProducto.Item(5).Text = 0
        'DESPERDICIO
                MskProducto.Item(6).Text = 0

End Sub

Public Sub BotonesVisibles4()
    If BanderaBotonesVisibles4 = True Then
        CmdBotones5.Item(0).Visible = True
        CmdBotones5.Item(1).Visible = True
        CmdBotones5.Item(2).Visible = True
        CmdBotones5.Item(3).Visible = True
        CmdBotones5.Item(4).Visible = True
    Else
        CmdBotones5.Item(0).Visible = False
        CmdBotones5.Item(1).Visible = False
        CmdBotones5.Item(2).Visible = False
        CmdBotones5.Item(3).Visible = False
        CmdBotones5.Item(4).Visible = False
    End If

End Sub

Public Sub Botones5()
    If Bandera5 = True Then
         FrameDetalleEmpleados.Enabled = True
         CmdBotones5.Item(0).Enabled = False
         CmdBotones5.Item(1).Enabled = True
         CmdBotones5.Item(2).Enabled = True
         CmdBotones5.Item(3).Enabled = False
         CmdBotones5.Item(4).Enabled = False
    Else
         FrameDetalleEmpleados.Enabled = False
         CmdBotones5.Item(0).Enabled = True
         CmdBotones5.Item(1).Enabled = False
         CmdBotones5.Item(2).Enabled = False
         CmdBotones5.Item(3).Enabled = True
         CmdBotones5.Item(4).Enabled = True
    End If

End Sub

Public Sub Limpia_CamposEmpleados()
        TxtTexto2.Item(9).Text = "" 'DOCUMENTO
        TxtTexto(15).Text = "" ' EMPLEADO
End Sub

Public Sub Llena_CamposEmpleados()
On Error Resume Next
    If RDetalleEmpleados.RecordCount > 0 Then
        'DOCUMENTO
            If IsNull(RDetalleEmpleados!Documento) Then
                TxtTexto2.Item(9).Text = ""
            Else
                TxtTexto2.Item(9).Text = RDetalleEmpleados!Documento
            End If
        'EMPLEADO
            If IsNull(RDetalleEmpleados!Empleado) Then
                TxtTexto.Item(15).Text = ""
            Else
                TxtTexto.Item(15).Text = RDetalleEmpleados!Empleado
            End If
    Else
            Limpia_CamposEmpleados
    End If
            
        If Err <> 0 Then
            
        End If

End Sub

Public Sub SumaParos()
On Error Resume Next
'SUMA LOS MINUTOS TIPO S
                        Set RTotalS = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalS, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'S'")
                            If RTotalS.RecordCount > 0 Then
                                If IsNull(RTotalS(0)) Then
                                    VTotalParoS = 0
                                Else
                                    VTotalParoS = RTotalS(0)
                                End If
                            Else
                                VTotalParoS = 0
                            End If
                        
                        'SUMA LOS MINUTOS TIPO N
                        Set RTotalN = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalN, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'N'")
                            If RTotalN.RecordCount > 0 Then
                                If IsNull(RTotalN(0)) Then
                                    VTotalParoN = 0
                                Else
                                    VTotalParoN = RTotalN(0)
                                End If
                            Else
                                VTotalParoN = 0
                            End If
                        
                        'SUMA LOS MINUTOS TIPO P
                        Set RTotalP = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo = 'P'")
                        
                            If RTotalP.RecordCount > 0 Then
                                If IsNull(RTotalP(0)) Then
                                    VTotalParoP = 0
                                Else
                                    VTotalParoP = RTotalP(0)
                                End If
                                
                            Else
                                VTotalParoP = 0
                            End If
                        
                        'SUMA LOS MINUTOS TIPO CF (CAMBIO DE FORMATO)
                        Set RTotalCF = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalCF, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo2 = 'CF'")
                            If RTotalCF.RecordCount > 0 Then
                                If IsNull(RTotalCF(0)) Then
                                    VTotalParoCF = 0
                                Else
                                    VTotalParoCF = RTotalCF(0)
                                End If
                            Else
                                VTotalParoCF = 0
                            End If
                            
                        'SUMA LOS MINUTOS TIPO MP (MANTENIMIENTO PREVENTIVO)
                        Set RTotalMP = New ADODB.Recordset
                        Call Abrir_Recordset(RTotalMP, "Select sum(DC.minutos) from DetalleCapturaParos DC, Paros P where DC.Documento = " & VDocumento & " And DC.Paro = P.CodigoParo And P.Tipo2 = 'MP'")
                            If RTotalMP.RecordCount > 0 Then
                                If IsNull(RTotalMP(0)) Then
                                    VTotalParoMP = 0
                                Else
                                    VTotalParoMP = RTotalMP(0)
                                End If
                            Else
                                VTotalParoMP = 0
                            End If
                        
                        Conexion.Execute "Update EncabezadoCapturaParos Set ParoS = " & VTotalParoS & ", ParoN = " & VTotalParoN & ", ParoP = " & VTotalParoP & ", ParoCF = " & VTotalParoCF & ", ParoMP = " & VTotalParoMP & " Where Documento = " & VDocumento
                        

                        
                        If Err <> 0 Then
                        End If
                    
End Sub

