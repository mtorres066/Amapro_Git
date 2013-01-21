VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CierreTarima 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre De Tarimas"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "CierreTarima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataCierreTarima 
      Caption         =   "Cierre De Tarima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CierreTarima"
      Top             =   7080
      Width           =   9135
   End
   Begin TabDlg.SSTab TabCierreTarima 
      Height          =   6135
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CierreTarima.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCierreTarima"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CierreTarima.frx":0624
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridCierreTarima"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "CierreTarima.frx":0A76
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DbGridCierreTarima 
         Bindings        =   "CierreTarima.frx":0EC8
         Height          =   5295
         Left            =   -74880
         OleObjectBlob   =   "CierreTarima.frx":0EE7
         TabIndex        =   18
         Top             =   720
         Width           =   8895
      End
      Begin VB.Frame FrameBusquedadeDatos 
         Caption         =   "Busqueda de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   -74880
         TabIndex        =   34
         Top             =   720
         Width           =   8775
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            MaxLength       =   2
            TabIndex        =   25
            Top             =   2640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha De Tarima Y Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   3
            Left            =   4440
            Picture         =   "CierreTarima.frx":2B2D
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Actual Y Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   3000
            Picture         =   "CierreTarima.frx":2E37
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha De Tarima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   1560
            Picture         =   "CierreTarima.frx":3141
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   7080
            TabIndex        =   24
            Top             =   2160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61603843
            CurrentDate     =   37441
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   5520
            TabIndex        =   23
            Top             =   2160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61603843
            CurrentDate     =   37441
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Actual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   120
            Picture         =   "CierreTarima.frx":3583
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label LblLin 
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
            Left            =   120
            TabIndex        =   51
            Top             =   3000
            Width           =   8415
         End
         Begin VB.Label LblLinea 
            Height          =   195
            Left            =   5040
            TabIndex        =   50
            Top             =   2640
            Width           =   405
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   14
            Left            =   5520
            TabIndex        =   46
            Top             =   1920
            Width           =   555
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   13
            Left            =   7080
            TabIndex        =   45
            Top             =   1920
            Width           =   510
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   6000
            TabIndex        =   27
            Top             =   4440
            Width           =   2535
            Caption         =   "Actualizar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CierreTarima.frx":43C5
            Accelerator     =   84
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   6
            Left            =   6000
            TabIndex        =   26
            Top             =   3720
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CierreTarima.frx":46DF
            Accelerator     =   83
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameCierreTarima 
         Caption         =   "Cierre De Tarima"
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
         Height          =   4575
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Liberados"
            DataSource      =   "DataCierreTarima"
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
            Index           =   8
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Saldo"
            DataSource      =   "DataCierreTarima"
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
            Index           =   7
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DataField       =   "Usuario"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   6
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3720
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Bodega"
            DataSource      =   "DataCierreTarima"
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
            Index           =   4
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Calidad"
            DataSource      =   "DataCierreTarima"
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
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Linea"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   4
            ToolTipText     =   "Busca Tarima"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Tarima"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   2
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   1
            Top             =   720
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Fecha"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   0
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "FechaProduccion"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Desperdicio"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   3
            Left            =   6840
            TabIndex        =   7
            Top             =   2880
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Descargar"
            DataSource      =   "DataCierreTarima"
            Height          =   285
            Index           =   5
            Left            =   1680
            TabIndex        =   8
            ToolTipText     =   "cantidad a descontar del saldo de tarima"
            Top             =   3360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3600
            Picture         =   "CierreTarima.frx":4B31
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   735
            Left            =   3480
            TabIndex        =   49
            Top             =   2880
            Width           =   15
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Left            =   3480
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label LblLabel 
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
            Index           =   2
            Left            =   2400
            TabIndex        =   44
            Top             =   2520
            Width           =   6135
         End
         Begin VB.Label LblLabel 
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
            Index           =   1
            Left            =   2400
            TabIndex        =   43
            Top             =   1800
            Width           =   6135
         End
         Begin VB.Label LblLabel 
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
            Index           =   0
            Left            =   3480
            TabIndex        =   42
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Desperdicio"
            Height          =   195
            Index           =   12
            Left            =   5880
            TabIndex        =   41
            Top             =   2880
            Width           =   840
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Liberados"
            Height          =   195
            Index           =   11
            Left            =   5880
            TabIndex        =   40
            Top             =   3240
            Width           =   690
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Total Descargar"
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
            Index           =   10
            Left            =   120
            TabIndex        =   39
            Top             =   3360
            Width           =   1380
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   38
            Top             =   3720
            Width           =   1065
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   37
            Top             =   4080
            Width           =   540
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Saldo en Bodega"
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
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   2880
            Width           =   1470
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Calidad"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   35
            Top             =   2160
            Width           =   525
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Bodega"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   33
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   390
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Tarima"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Actual"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   945
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   6360
      Width           =   1695
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "2999;1085"
      Picture         =   "CierreTarima.frx":4F73
      Accelerator     =   65
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   2
      Left            =   2040
      TabIndex        =   12
      Top             =   6360
      Width           =   1695
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "2999;1085"
      Picture         =   "CierreTarima.frx":54B5
      Accelerator     =   71
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   3
      Left            =   3840
      TabIndex        =   13
      Top             =   6360
      Width           =   1695
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "2999;1085"
      Accelerator     =   67
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   4
      Left            =   5640
      TabIndex        =   14
      ToolTipText     =   "si borra, lo descargado regresa al saldo de la tarima"
      Top             =   6360
      Width           =   1695
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "2999;1085"
      Accelerator     =   66
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   5
      Left            =   7440
      TabIndex        =   15
      Top             =   6360
      Width           =   1695
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "2999;1085"
      Picture         =   "CierreTarima.frx":59F7
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "CierreTarima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim RBuscaFichaTecnica As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaTarima As Recordset

Dim VDescargar As Long

Dim VFichaTecnica As String
Dim VTarima As Long
Dim VFechaProduccion As Date
Dim VLinea As String

Dim VUltimaFichaTecnica As String
Dim VUltimaTarima As Long
Dim VUltimaFechaProduccion As String
Dim VUltimaLinea As String
Dim VUltimaDescargar As Long
Dim VUltimaObservaciones As String


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataCierreTarima.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                        Msk.Item(0).Text = Date
                        Msk.Item(0).SetFocus
                        Txttexto.Item(0).Text = VUltimaFichaTecnica
                        Txttexto.Item(1).Text = Val(VUltimaTarima) + 1
                        Msk.Item(1).Text = VUltimaFechaProduccion
                        Txttexto.Item(2).Text = VUltimaLinea
                        Txttexto.Item(5).Text = VUltimaObservaciones
                        Msk.Item(5).Text = VUltimaDescargar
                        Txttexto.Item(6).Text = GUsuario
        'EDITAR
        'ElseIf Index = 1 Then
        '                .Edit
        '                If Err.Number > 0 Then
        '                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
        '                        Exit Sub
        '                End If
        '        Bandera = True
        '        botones
        '                Msk.Item(0).SetFocus
        '
        'GRABAR
        ElseIf Index = 2 Then
        
                'ASIGNA VALORES A LAS VARIABLES PARA DESPUES DE GRABAR QUEDEN GUARDADOS LOS VALORES INICIALES
                VDescargar = Msk.Item(5).Text
                VFichaTecnica = Txttexto.Item(0).Text
                VTarima = Txttexto.Item(1).Text
                VFechaProduccion = Msk.Item(1).Text
                VLinea = Txttexto.Item(2).Text
                
                'GUARDA VARIABLES PARA DESPLEGARLAS CUANDO AGREGE
                VUltimaFichaTecnica = Txttexto.Item(0).Text
                VUltimaTarima = Txttexto.Item(1).Text
                VUltimaFechaProduccion = Msk.Item(1).Text
                VUltimaLinea = Txttexto.Item(2).Text
                VUltimaObservaciones = Txttexto.Item(5).Text
                VUltimaDescargar = Msk.Item(5).Text
                
                
                'REVISA LA FECHA
                If Not IsDate(Msk.Item(0).Text) Then
                    MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                    Msk.Item(0).SetFocus
                    Exit Sub
                End If
                
                'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
                Set RBuscaTarima = Db.OpenRecordset("Select Saldo, Calidad From DetalleEntradasProductoTermina Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                    If RBuscaTarima.RecordCount > 0 Then
                    Else
                        MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "No Se Puede Grabar"
                        Txttexto.Item(0).SetFocus
                        Exit Sub
                    End If
                    
                'REVISA LA CANTIDAD A DESCONTAR NO PUEDE SER MAYOR QUE EL SALDO
                If Val(Msk.Item(5).Text) > Val(Txttexto.Item(7).Text) Then
                    MsgBox "La Cantidad A Descontar No Puede Ser Mayor Que El Saldo De La Tarima", vbOKOnly + vbInformation, "Informacion"
                    Msk.Item(5).SetFocus
                    Exit Sub
                End If
                    
                'GRABA DATOS
                .Update
                
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                        
                'BUSCA EL SALDO DE LA TARIMA
                Set RBuscaTarima = Db.OpenRecordset("Select Saldo, Salidas From DetalleEntradasProductoTermina Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                    If RBuscaTarima.RecordCount > 0 Then
                        RBuscaTarima.Edit
                            'LE RESTA AL SALDO DE LA TARIMA LA CANTIDAD A DESCONTAR
                            RBuscaTarima!Saldo = Val(RBuscaTarima!Saldo) - VDescargar
                            RBuscaTarima!Salidas = Val(RBuscaTarima!Salidas) + VDescargar
                        RBuscaTarima.Update
                    Else
                        MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "No Se Puede Grabar"
                        Txttexto.Item(0).SetFocus
                        Exit Sub
                    End If
                
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 3 Then
                .CancelUpdate
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
        
        'VERIFICA SI ESTA AUTORIZADO PARA BORRAR
        If GBorrar = False Then
               MsgBox "Usted No Tiene Acceso a Esta Funcion, Llame Al Supervisor", vbOKOnly + vbInformation, "Informacion"
               Exit Sub
        End If
                
                'ASIGNA VALORES A LAS VARIABLES PARA DESPUES DE GRABAR QUEDEN GUARDADOS LOS VALORES INICIALES
                VDescargar = Msk.Item(5).Text
                VFichaTecnica = Txttexto.Item(0).Text
                VTarima = Txttexto.Item(1).Text
                VFechaProduccion = Msk.Item(1).Text
                VLinea = Txttexto.Item(2).Text
        
                VMensaje = MsgBox("Esta Seguro De Borrar El Registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                
                If VMensaje = vbYes Then
                    .Delete
                    .MoveNext
                            If Err.Number > 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                        
                    'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
                    Set RBuscaTarima = Db.OpenRecordset("Select Saldo, Salidas From DetalleEntradasProductoTermina Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                        If RBuscaTarima.RecordCount > 0 Then
                            RBuscaTarima.Edit
                                'LE SUMA AL SALDO DE LA TARIMA LA CANTIDAD A DESCONTAR
                                RBuscaTarima!Saldo = Val(RBuscaTarima!Saldo) + VDescargar
                                RBuscaTarima!Salidas = Val(RBuscaTarima!Salidas) - VDescargar
                            RBuscaTarima.Update
                        Else
                            MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "Informacion                        "
                        End If
                    End If
                
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataCierreTarima.RecordSource = ("Select * From CierreTarima Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fecha")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataCierreTarima.RecordSource = ("Select * From CierreTarima Where FechaProduccion >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaProduccion <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fecha")
                    ElseIf OptBusqueda.Item(2).Value = True Then
                        DataCierreTarima.RecordSource = ("Select * From CierreTarima Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & Txtlin.Text & "' Order by Fecha")
                    ElseIf OptBusqueda.Item(3).Value = True Then
                        DataCierreTarima.RecordSource = ("Select * From CierreTarima Where FechaProduccion >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaProduccion <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & Txtlin.Text & "' Order by Fecha")
                    End If
                    DataCierreTarima.Refresh
                    DBGridCierreTarima.Refresh
                    TabCierreTarima.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataCierreTarima.RecordSource = "Select * From CierreTarima"
                    DataCierreTarima.Refresh
                    DBGridCierreTarima.Refresh
                    TabCierreTarima.Tab = 1
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameCierreTarima.Enabled = True
         DataCierreTarima.Visible = False
         DBGridCierreTarima.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameCierreTarima.Enabled = False
         DataCierreTarima.Visible = True
         DBGridCierreTarima.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub

Private Sub DbGridCierreTarima_HeadClick(ByVal ColIndex As Integer)
    DataCierreTarima.RecordSource = ("Select * from CierreTarima order by " & DBGridCierreTarima.Columns(ColIndex).DataField)
    DataCierreTarima.Refresh
    DBGridCierreTarima.Refresh

End Sub

Private Sub Form_Load()
    DataCierreTarima.ConnectionString = GTipoProveedor
    DataCierreTarima.Refresh
End Sub



Private Sub Msk_GotFocus(Index As Integer)
        Msk.Item(Index).SelStart = 0
        Msk.Item(Index).SelLength = Len(Msk.Item(Index))
End Sub

Private Sub Msk_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Msk_LostFocus(Index As Integer)
        If Index = 3 Then
            'AL SALDO DE LA TARIMA LE RESTA EL DESPERDICIO Y ESE ES EL PRODUCTO LIBERADO
            Txttexto.Item(8).Text = Val(Txttexto.Item(7).Text) - Val(Msk.Item(3).Text)
        End If
End Sub


Private Sub OptBusqueda_Click(Index As Integer)
    If (Index = 2 Or Index = 3) Then
        LblLin.Caption = "Linea"
        Txtlin.Visible = True
        LblLin.Caption = ""
    Else
        LblLin.Caption = ""
        Txtlin.Visible = False
        LblLin.Caption = ""
    End If
End Sub

Private Sub TabCierreTarima_Click(PreviousTab As Integer)
        If TabCierreTarima.Tab = 2 Then
                DTPFecIni.Value = Date
                DTPFecFin.Value = Date
        End If
End Sub


Private Sub TxtLin_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                    LblLin.Caption = RBuscaLinea!Descrip
            Else
                    LblLin.Caption = ""
            End If

End Sub

Private Sub TxtTexto_Change(Index As Integer)
    'BUSCA FICHA TECNICA
    If Index = 0 Then
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & Txttexto.Item(0).Text & "'")
            If RBuscaFichaTecnica.RecordCount > 0 Then
                    Lbllabel.Item(0).Caption = RBuscaFichaTecnica!Descrip
            Else
                    Lbllabel.Item(0).Caption = ""
            End If
    'BUSCA LINEA
    ElseIf Index = 2 Then
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txttexto.Item(2).Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                    Lbllabel.Item(1).Caption = RBuscaLinea!Descrip
            Else
                    Lbllabel.Item(1).Caption = ""
            End If
    'BUSCA BODEGA DE PRODUCTO TERMINADO
    ElseIf Index = 4 Then
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & Txttexto.Item(4).Text & "'")
            If RBuscaBodega.RecordCount > 0 Then
                    Lbllabel.Item(2).Caption = RBuscaBodega!Descripcion
            Else
                    Lbllabel.Item(2).Caption = ""
            End If
    End If
End Sub


Private Sub TxtTexto_GotFocus(Index As Integer)
    Txttexto.Item(Index).SelStart = 0
    Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub


Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
'On Error Resume Next
    If Index = 2 Then
    
        'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
        Set RBuscaTarima = Db.OpenRecordset("Select Saldo, Calidad, Bodega From DetalleEntradasProductoTermina Where FechaProduccion = #" & Format(Msk.Item(1).Text, "mm/dd/yyyy") & "# And Tarima = " & Txttexto.Item(1).Text & " And Linea = '" & Txttexto.Item(2).Text & "' And FichaTecnica = '" & Txttexto.Item(0).Text & "'")
            If RBuscaTarima.RecordCount > 0 Then
                Txttexto.Item(7).Text = RBuscaTarima!Saldo
                Txttexto.Item(4).Text = RBuscaTarima!Bodega
                Txttexto.Item(3).Text = RBuscaTarima!Calidad
            Else
                MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "Informacion"
            End If
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
            End If
            
    End If
        
End Sub
