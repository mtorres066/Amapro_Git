VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form TarimasProductoTerminado 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarimas De Producto Terminado"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "TarimasProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataTarimas 
      Caption         =   "Tarimas De Producto Terminado"
      Connect         =   "Access"
      DatabaseName    =   "C:\erick\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   320
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasProductoTerminado"
      Top             =   7800
      Width           =   11655
   End
   Begin TabDlg.SSTab TabBultos 
      Height          =   7215
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "TarimasProductoTerminado.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBultos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "TarimasProductoTerminado.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridTarimas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "TarimasProductoTerminado.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridTarimas 
         Bindings        =   "TarimasProductoTerminado.frx":1488
         Height          =   6015
         Left            =   -74880
         OleObjectBlob   =   "TarimasProductoTerminado.frx":14A2
         TabIndex        =   23
         Top             =   720
         Width           =   11655
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
         Height          =   6375
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   11655
         Begin VB.Frame FrameFechas 
            Height          =   855
            Left            =   6720
            TabIndex        =   56
            Top             =   2640
            Width           =   4215
            Begin MSComCtl2.DTPicker DtpFecIni 
               Height          =   255
               Left            =   720
               TabIndex        =   57
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   23986179
               CurrentDate     =   37455
            End
            Begin MSComCtl2.DTPicker DtpFecFin 
               Height          =   255
               Left            =   2760
               TabIndex        =   58
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   23986179
               CurrentDate     =   37455
            End
            Begin VB.Label lblFieldLabel 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hasta"
               Height          =   195
               Index           =   15
               Left            =   2280
               TabIndex        =   60
               Top             =   360
               Width           =   420
            End
            Begin VB.Label lblFieldLabel 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Desde"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   59
               Top             =   360
               Width           =   465
            End
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Batch Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   5
            Left            =   7320
            Picture         =   "TarimasProductoTerminado.frx":375E
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   360
            Width           =   1300
         End
         Begin VB.TextBox TxtBusqueda2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   26
            Top             =   3600
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   1335
            Index           =   4
            Left            =   5880
            Picture         =   "TarimasProductoTerminado.frx":3A68
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas Y Ficha Tecnica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   3
            Left            =   4440
            Picture         =   "TarimasProductoTerminado.frx":4332
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas Y Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   2
            Left            =   3000
            Picture         =   "TarimasProductoTerminado.frx":4BFC
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   360
            Width           =   1300
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   27
            Top             =   3960
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   1
            Left            =   1560
            Picture         =   "TarimasProductoTerminado.frx":4F0E
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   1335
            Index           =   0
            Left            =   120
            Picture         =   "TarimasProductoTerminado.frx":5218
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1300
         End
         Begin VB.Label LblBusqueda2 
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
            Left            =   5160
            TabIndex        =   53
            Top             =   3600
            Width           =   3135
         End
         Begin VB.Label LblBusqueda 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   37
            Top             =   3960
            Width           =   3135
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   8400
            TabIndex        =   29
            Top             =   5160
            Width           =   2535
            Caption         =   "Seleccionar Todos"
            PicturePosition =   196613
            Size            =   "4471;1085"
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
            Left            =   8400
            TabIndex        =   28
            Top             =   4440
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameBultos 
         Caption         =   "Datos De La Tarima"
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
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Batch"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   15
            Left            =   1800
            TabIndex        =   5
            Top             =   2280
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   4
            Top             =   1920
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Casilla"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   9
            Left            =   1800
            MaxLength       =   5
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   4080
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Pasillo"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   8
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   9
            Top             =   3720
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Bodega"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   7
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   8
            Top             =   3360
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Calidad"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   6
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   7
            Top             =   3000
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Batch"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   5
            Left            =   1800
            TabIndex        =   6
            Top             =   2640
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Tarima"
            DataSource      =   "DataTarimas"
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
            Left            =   1800
            TabIndex        =   2
            Top             =   1200
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Usuario"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   14
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   15
            Top             =   5880
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Saldo"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   13
            Left            =   1800
            TabIndex        =   14
            Top             =   5520
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Salidas"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   12
            Left            =   1800
            TabIndex        =   13
            Top             =   5160
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   11
            Left            =   1800
            TabIndex        =   12
            Top             =   4800
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Bin"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   11
            Top             =   4440
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FichaTecnica"
            DataSource      =   "DataTarimas"
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
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   1
            Top             =   840
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FechaProduccion"
            DataSource      =   "DataTarimas"
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   3
            Top             =   1560
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataTarimas"
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
            Left            =   1800
            TabIndex        =   0
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Orden"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   61
            Top             =   2280
            Width           =   435
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
            Left            =   3840
            TabIndex        =   55
            Top             =   3360
            Width           =   7095
         End
         Begin VB.Label LblFicTec 
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
            TabIndex        =   49
            Top             =   840
            Width           =   7095
         End
         Begin VB.Label LblLin 
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
            TabIndex        =   48
            Top             =   1920
            Width           =   7095
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Linea"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   47
            Top             =   1920
            Width           =   390
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   46
            Top             =   5880
            Width           =   540
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tarima"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   45
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Batch"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   420
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Calidad"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   43
            Top             =   3000
            Width           =   525
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bodega"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   42
            Top             =   3360
            Width           =   555
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pasillo"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   41
            Top             =   3720
            Width           =   450
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Casilla"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   40
            Top             =   4080
            Width           =   450
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Saldo"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   5520
            Width           =   405
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inicio"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   38
            Top             =   4800
            Width           =   375
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Salidas"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   35
            Top             =   5160
            Width           =   510
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bin"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   34
            Top             =   4440
            Width           =   225
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Produccion"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Top             =   1560
            Width           =   1305
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Documento"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   825
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   1800
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Accelerator     =   65
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   18
      Top             =   7320
      Width           =   1800
      Caption         =   "Editar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "TarimasProductoTerminado.frx":565A
      Accelerator     =   69
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   19
      Top             =   7320
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "TarimasProductoTerminado.frx":5B9C
      Accelerator     =   71
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   20
      Top             =   7320
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "TarimasProductoTerminado.frx":60DE
      Accelerator     =   67
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   21
      Top             =   7320
      Width           =   1800
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "TarimasProductoTerminado.frx":6620
      Accelerator     =   66
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   5
      Left            =   9720
      TabIndex        =   22
      Top             =   7320
      Width           =   1800
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "TarimasProductoTerminado.frx":6B62
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "TarimasProductoTerminado"
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
Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataTarimas.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
        'EDITAR
        ElseIf Index = 1 Then
                        .Edit
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
        'GRABAR
        ElseIf Index = 2 Then
                    
                .Update
                
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
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
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If VMensaje = vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number > 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where Documento = '" & TxtBusqueda.Text & "' Order By Batch, Tarima")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where FechaProduccion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaProduccion <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Order By Batch, Tarima")
                    ElseIf OptBusqueda.Item(2).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where FechaProduccion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaProduccion <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBusqueda.Text & "' Order By Batch, Tarima")
                    ElseIf OptBusqueda.Item(3).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where FechaProduccion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaProduccion <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And FichaTecnica = '" & TxtBusqueda.Text & "' Order By Batch, Tarima")
                    ElseIf OptBusqueda.Item(4).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where Batch = " & TxtBusqueda.Text & " Order By Batch, Tarima")
                    ElseIf OptBusqueda.Item(5).Value = True Then
                        DataTarimas.RecordSource = ("Select * From DetalleEntradasProductoTermina where Batch = " & TxtBusqueda.Text & " And Linea = '" & TxtBusqueda2.Text & "' Order By Batch, Tarima")
                    End If
                    DataTarimas.Refresh
                    DGridTarimas.Refresh
                    TabBultos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataTarimas.RecordSource = "Select * From DetalleEntradasProductoTermina"
                    DataTarimas.Refresh
                    DGridTarimas.Refresh
                    TabBultos.Tab = 1
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameBultos.Enabled = True
         DataTarimas.Visible = False
         DGridTarimas.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameBultos.Enabled = False
         DataTarimas.Visible = True
         DGridTarimas.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub DGridTarimas_HeadClick(ByVal ColIndex As Integer)
        DataTarimas.RecordSource = "Select * From DetalleEntradasProductoTermina Order By " & DGridTarimas.Columns(ColIndex).DataField
        DataTarimas.Refresh
        DGridTarimas.Refresh
End Sub

Private Sub Form_Load()
            DataTarimas.ConnectionString = GTipoProveedor
            DataTarimas.Refresh
End Sub



Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Documento"
            LblBusqueda2.Caption = ""
            TxtBusqueda.Visible = True
            TxtBusqueda2.Visible = False
            FrameFechas.Visible = False
            TxtBusqueda.SetFocus
    ElseIf Index = 1 Then
            LblBusqueda.Caption = ""
            LblBusqueda2.Caption = ""
            TxtBusqueda.Visible = False
            TxtBusqueda2.Visible = False
            FrameFechas.Visible = True
    ElseIf Index = 2 Then
            LblBusqueda.Caption = "Linea"
            LblBusqueda2.Caption = ""
            TxtBusqueda.Visible = True
            TxtBusqueda2.Visible = False
            FrameFechas.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 3 Then
            LblBusqueda.Caption = "Ficha Tecnica"
            LblBusqueda2.Caption = ""
            TxtBusqueda.Visible = True
            TxtBusqueda2.Visible = False
            FrameFechas.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 4 Then
            LblBusqueda.Caption = "Batch"
            LblBusqueda2.Caption = ""
            TxtBusqueda.Visible = True
            TxtBusqueda2.Visible = False
            FrameFechas.Visible = False
            TxtBusqueda.SetFocus
    ElseIf Index = 5 Then
            LblBusqueda.Caption = "Batch"
            LblBusqueda2.Caption = "Linea"
            TxtBusqueda.Visible = True
            TxtBusqueda2.Visible = True
            FrameFechas.Visible = False
            TxtBusqueda.SetFocus
    End If

            
End Sub

Private Sub TabBultos_Click(PreviousTab As Integer)
        If TabBultos.Tab = 2 Then
                DtpFecIni.Value = Date
                DtpFecFin.Value = Date
        End If
        
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

Private Sub TxtBusqueda2_GotFocus()
        TxtBusqueda2.SelStart = 0
        TxtBusqueda2.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub TxtBusqueda2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        'FICHA TECNICA
        If Index = 1 Then
            Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFicTec.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFicTec.Caption = ""
                End If
        'LINEA
        ElseIf Index = 4 Then
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(4).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLin.Caption = RBuscaLinea!Descrip
                Else
                    LblLin.Caption = ""
                End If
        'BODEGA
        ElseIf Index = 7 Then
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtTexto.Item(7).Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBod.Caption = RBuscaBodega!Descripcion
                Else
                    LblBod.Caption = ""
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
