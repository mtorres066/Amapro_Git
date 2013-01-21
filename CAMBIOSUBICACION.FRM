VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CambiosUbicacion 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios De Ubicacion De Tarimas"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "CambiosUbicacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataCambiosUbicacion 
      Caption         =   "Cambios Ubicacion"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CambiosUbicacion"
      Top             =   8160
      Width           =   9135
   End
   Begin TabDlg.SSTab TabCambiosUbicacion 
      Height          =   7335
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CambiosUbicacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCambiosUbicacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CambiosUbicacion.frx":0624
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridCambiosUbicacion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "CambiosUbicacion.frx":0A76
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DBGridCambiosUbicacion 
         Bindings        =   "CambiosUbicacion.frx":0EC8
         Height          =   6495
         Left            =   -74880
         OleObjectBlob   =   "CambiosUbicacion.frx":0EEB
         TabIndex        =   22
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
         Height          =   6495
         Left            =   -74880
         TabIndex        =   34
         Top             =   720
         Width           =   8775
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   5640
            TabIndex        =   25
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24379395
            CurrentDate     =   37441
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   4080
            TabIndex        =   24
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24379395
            CurrentDate     =   37441
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   975
            Index           =   0
            Left            =   360
            Picture         =   "CambiosUbicacion.frx":2B3D
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
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
            Left            =   4080
            TabIndex        =   41
            Top             =   480
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
            Left            =   5640
            TabIndex        =   40
            Top             =   480
            Width           =   510
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   6000
            TabIndex        =   27
            Top             =   5640
            Width           =   2535
            Caption         =   "Actualizar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CambiosUbicacion.frx":2E47
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
            Top             =   4920
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CambiosUbicacion.frx":3161
            Accelerator     =   83
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameCambiosUbicacion 
         Caption         =   "Cambios De Ubicacion"
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
         Height          =   6375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NuevaBin"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   11
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   13
            Top             =   5880
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NuevaCasilla"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   10
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   12
            Top             =   5520
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NuevaPasillo"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   9
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   11
            Top             =   5160
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Bin"
            DataSource      =   "DataCambiosUbicacion"
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
            MaxLength       =   5
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Casilla"
            DataSource      =   "DataCambiosUbicacion"
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
            Index           =   6
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   3480
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Pasillo"
            DataSource      =   "DataCambiosUbicacion"
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
            Index           =   5
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NuevaBodega"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   8
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   10
            Top             =   4800
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DataField       =   "Usuario"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   12
            Left            =   6840
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Bodega"
            DataSource      =   "DataCambiosUbicacion"
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
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Calidad"
            DataSource      =   "DataCambiosUbicacion"
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
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            DataField       =   "Linea"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   4
            ToolTipText     =   "Busca Tarima"
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            DataField       =   "Tarima"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   2
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataCambiosUbicacion"
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   1
            Top             =   720
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Fecha"
            DataSource      =   "DataCambiosUbicacion"
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
            DataSource      =   "DataCambiosUbicacion"
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
            BackColor       =   12640511
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ubicacion Nueva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   18
            Left            =   6240
            TabIndex        =   51
            Top             =   4320
            Width           =   2025
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ubicacion Anterior"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   6240
            TabIndex        =   50
            Top             =   2280
            Width           =   2235
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nuevo Bin"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   49
            Top             =   5880
            Width           =   750
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nueva Casilla"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   48
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nuevo Pasillo"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   47
            Top             =   5160
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Bin"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   46
            Top             =   3840
            Width           =   225
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Casilla"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   45
            Top             =   3480
            Width           =   450
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Pasillo"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   44
            Top             =   3120
            Width           =   450
         End
         Begin VB.Label LblLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   43
            Top             =   4800
            Width           =   5775
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nueva Bodega"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   42
            Top             =   4800
            Width           =   1080
         End
         Begin VB.Label LblLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   39
            Top             =   2760
            Width           =   5775
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
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   38
            Top             =   1800
            Width           =   5895
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
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   37
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   8
            Left            =   6240
            TabIndex        =   36
            Top             =   240
            Width           =   540
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
            BackColor       =   &H0080C0FF&
            Caption         =   "Bodega"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   33
            Top             =   2760
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
            Caption         =   "Fecha Produccion"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   1305
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
            Caption         =   "Fecha"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   450
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            Height          =   1575
            Index           =   0
            Left            =   120
            Top             =   2640
            Width           =   8415
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   1575
            Index           =   1
            Left            =   120
            Top             =   4680
            Width           =   8415
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   495
      Index           =   4
      Left            =   5520
      TabIndex        =   18
      Top             =   7560
      Width           =   1695
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "2999;873"
      Picture         =   "CambiosUbicacion.frx":35B3
      Accelerator     =   98
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   7560
      Width           =   1700
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "2999;873"
      Picture         =   "CambiosUbicacion.frx":3AF5
      Accelerator     =   65
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   16
      Top             =   7560
      Width           =   1695
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "2999;873"
      Picture         =   "CambiosUbicacion.frx":4037
      Accelerator     =   71
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   17
      Top             =   7560
      Width           =   1695
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "2999;873"
      Picture         =   "CambiosUbicacion.frx":4579
      Accelerator     =   67
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   495
      Index           =   5
      Left            =   7320
      TabIndex        =   19
      Top             =   7560
      Width           =   1695
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "2999;873"
      Picture         =   "CambiosUbicacion.frx":4ABB
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "CambiosUbicacion"
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

Dim VFichaTecnica As String
Dim VTarima As Long
Dim VFechaProduccion As Date
Dim VLinea As String

Dim VBodegaNueva As String
Dim VPasilloNueva As String
Dim VCasillaNueva As String
Dim VBinNueva As String




Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataCambiosUbicacion.Recordset
    
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
                        TxtTexto.Item(12).Text = GUsuario
        
        'GRABAR
        ElseIf Index = 2 Then
        
                'ASIGNA VALORES A VARIABLES PARA BUSCAR DATOS DESPUES DE GRABAR
                VFichaTecnica = TxtTexto.Item(0).Text
                VTarima = TxtTexto.Item(1).Text
                VFechaProduccion = Msk.Item(1).Text
                VLinea = TxtTexto.Item(2).Text
        
                'ASIGNA VALORES A LAS VARIABLES PARA LA NUEVA UBICACION
                VBodegaNueva = TxtTexto.Item(8).Text
                VPasilloNueva = TxtTexto.Item(9).Text
                VCasillaNueva = TxtTexto.Item(10).Text
                VBinNueva = TxtTexto.Item(11).Text
        
                'REVISA LA FECHA
                If Not IsDate(Msk.Item(0).Text) Then
                    MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                    Msk.Item(0).SetFocus
                    Exit Sub
                End If
                
                'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
                Set RBuscaTarima = Db.OpenRecordset("Select CalidaD, Bodega, Pasillo, Casilla, Bin From DetalleEntradasProductoTerminado Where FechaProduccion = #" & Format(Msk.Item(1).Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(1).Text & " And Linea = '" & TxtTexto.Item(2).Text & "' And FichaTecnica = '" & TxtTexto.Item(0).Text & "'")
                    If RBuscaTarima.RecordCount > 0 Then
                    Else
                        MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "No Se Puede Grabar"
                        TxtTexto.Item(0).SetFocus
                        Exit Sub
                    End If
                   
                    
                'GRABA DATOS
                .Update
                
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                        
                'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
                Set RBuscaTarima = Db.OpenRecordset("Select Bodega, Pasillo, Casilla, Bin From DetalleEntradasProductoTerminado Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                    If RBuscaTarima.RecordCount > 0 Then
                        RBuscaTarima.Edit
                            'CAMBIA LA UBICACION A LA TARIMA
                            RBuscaTarima!Bodega = VBodegaNueva
                            RBuscaTarima!Pasillo = VPasilloNueva
                            RBuscaTarima!Casilla = VCasillaNueva
                            RBuscaTarima!Bin = VBinNueva
                        RBuscaTarima.Update
                    Else
                        MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "No Se Puede Grabar"
                        TxtTexto.Item(0).SetFocus
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
                'ASIGNA VALORES A LAS VARIABLES PARA LA NUEVA UBICACION
                VBodegaNueva = TxtTexto.Item(4).Text
                VPasilloNueva = TxtTexto.Item(5).Text
                VCasillaNueva = TxtTexto.Item(6).Text
                VBinNueva = TxtTexto.Item(7).Text
        
                VMensaje = MsgBox("Esta Seguro De Borrar El Registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                
                If vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number > 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
                
                'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
                Set RBuscaTarima = Db.OpenRecordset("Select Bodega, Pasillo, Casilla, Bin From DetalleEntradasProductoTerminado Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                    If RBuscaTarima.RecordCount > 0 Then
                        RBuscaTarima.Edit
                            'CAMBIA LA UBICACION A LA TARIMA
                            RBuscaTarima!Bodega = VBodegaNueva
                            RBuscaTarima!Pasillo = VPasilloNueva
                            RBuscaTarima!Casilla = VCasillaNueva
                            RBuscaTarima!Bin = VBinNueva
                        RBuscaTarima.Update
                    End If
                               
                
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataCambiosUbicacion.RecordSource = ("Select * From CambiosUbicacion Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Order by Fecha")
                    End If
                    DataCambiosUbicacion.Refresh
                    DBGridCambiosUbicacion.Refresh
                    TabCambiosUbicacion.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataCambiosUbicacion.RecordSource = "Select * From CambiosUbicacion"
                    DataCambiosUbicacion.Refresh
                    DBGridCambiosUbicacion.Refresh
                    TabCambiosUbicacion.Tab = 1
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
         FrameCambiosUbicacion.Enabled = True
         DataCambiosUbicacion.Visible = False
         DBGridCambiosUbicacion.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameCambiosUbicacion.Enabled = False
         DataCambiosUbicacion.Visible = True
         DBGridCambiosUbicacion.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub



Private Sub DBGridCambiosUbicacion_HeadClick(ByVal ColIndex As Integer)
    DataCambiosUbicacion.RecordSource = ("Select * from CambiosUbicacion order by " & DBGridCambiosUbicacion.Columns(ColIndex).DataField)
    DataCambiosUbicacion.Refresh
    DBGridCambiosUbicacion.Refresh
End Sub

Private Sub Form_Load()
    DataCambiosUbicacion.DatabaseName = BasedeDatos
    
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
            TxtTexto.Item(8).Text = Val(TxtTexto.Item(7).Text) - Val(Msk.Item(3).Text)
        End If
End Sub


Private Sub TabCambiosubicacion_Click(PreviousTab As Integer)
        If TabCambiosUbicacion.Tab = 2 Then
                DtpFecIni.Value = Date
                DtpFecFin.Value = Date
        End If
End Sub


Private Sub TxtTexto_Change(Index As Integer)
    'BUSCA FICHA TECNICA
    If Index = 0 Then
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
            If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblLabel.Item(0).Caption = RBuscaFichaTecnica!Descrip
            Else
                    LblLabel.Item(0).Caption = ""
            End If
    'BUSCA LINEA
    ElseIf Index = 2 Then
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(2).Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                    LblLabel.Item(1).Caption = RBuscaLinea!Descrip
            Else
                    LblLabel.Item(1).Caption = ""
            End If
    'BUSCA BODEGA DE PRODUCTO TERMINADO
    ElseIf Index = 4 Then
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtTexto.Item(4).Text & "'")
            If RBuscaBodega.RecordCount > 0 Then
                    LblLabel.Item(2).Caption = RBuscaBodega!Descripcion
            Else
                    LblLabel.Item(2).Caption = ""
            End If
    'BUSCA BODEGA DE PRODUCTO TERMINADO
    ElseIf Index = 8 Then
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtTexto.Item(8).Text & "'")
            If RBuscaBodega.RecordCount > 0 Then
                    LblLabel.Item(3).Caption = RBuscaBodega!Descripcion
            Else
                    LblLabel.Item(3).Caption = ""
            End If
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
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
    If Index = 2 Then
    
        'BUSCA SI EXISTE LA TARIMA EN LAS ENTRADAS X EVADEVA DE PRODUCTO TERMINADO(INVENTARIO)
        Set RBuscaTarima = Db.OpenRecordset("Select Calidad, Bodega, Pasillo, Casilla, Bin From DetalleEntradasProductoTerminado Where FechaProduccion = #" & Format(Msk.Item(1).Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(1).Text & " And Linea = '" & TxtTexto.Item(2).Text & "' And FichaTecnica = '" & TxtTexto.Item(0).Text & "'")
            If RBuscaTarima.RecordCount > 0 Then
                TxtTexto.Item(3).Text = RBuscaTarima!Calidad
                TxtTexto.Item(4).Text = RBuscaTarima!Bodega
                If Not IsNull(RBuscaTarima!Pasillo) Then
                    TxtTexto.Item(5).Text = RBuscaTarima!Pasillo
                End If
                If Not IsNull(RBuscaTarima!Casilla) Then
                    TxtTexto.Item(6).Text = RBuscaTarima!Casilla
                End If
                If Not IsNull(RBuscaTarima!Bin) Then
                    TxtTexto.Item(7).Text = RBuscaTarima!Bin
                End If
            Else
                MsgBox "Tarima No Existe, En Entradas De Producto Terminado", vbOKOnly + vbInformation, "Informacion"
            End If
    End If
        
End Sub
