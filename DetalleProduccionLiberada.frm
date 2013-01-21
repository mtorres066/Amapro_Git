VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DetalleProduccionLiberada 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Detalle Produccion Liberada"
   ClientHeight    =   6732
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   8748
   Icon            =   "DetalleProduccionLiberada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6732
   ScaleWidth      =   8748
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framebuscar 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Data DataBuscar 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   3840
         TabIndex        =   36
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   38
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "DetalleProduccionLiberada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "DetalleProduccionLiberada.frx":237C
         Height          =   5415
         Left            =   120
         OleObjectBlob   =   "DetalleProduccionLiberada.frx":2395
         TabIndex        =   33
         Top             =   1080
         Width           =   8175
      End
      Begin VB.Label LblBusqueda 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   600
      Index           =   5
      Left            =   6840
      MouseIcon       =   "DetalleProduccionLiberada.frx":2D6D
      Picture         =   "DetalleProduccionLiberada.frx":31AF
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   600
      Index           =   4
      Left            =   5160
      MouseIcon       =   "DetalleProduccionLiberada.frx":5221
      Picture         =   "DetalleProduccionLiberada.frx":5663
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   600
      Index           =   2
      Left            =   1800
      MouseIcon       =   "DetalleProduccionLiberada.frx":5B95
      Picture         =   "DetalleProduccionLiberada.frx":5FD7
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   600
      Index           =   1
      Left            =   360
      MouseIcon       =   "DetalleProduccionLiberada.frx":6509
      MousePointer    =   99  'Custom
      Picture         =   "DetalleProduccionLiberada.frx":694B
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " "
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   600
      Index           =   0
      Left            =   120
      MouseIcon       =   "DetalleProduccionLiberada.frx":6E7D
      Picture         =   "DetalleProduccionLiberada.frx":72BF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   600
      Index           =   3
      Left            =   3480
      MouseIcon       =   "DetalleProduccionLiberada.frx":77F1
      Picture         =   "DetalleProduccionLiberada.frx":7C33
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   1600
   End
   Begin TabDlg.SSTab TabParos 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14838
      _ExtentY        =   9335
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "DetalleProduccionLiberada.frx":8165
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameParos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "DetalleProduccionLiberada.frx":847F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridParos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "DetalleProduccionLiberada.frx":8799
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusqueda"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameBusqueda 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   18
         Top             =   720
         Width           =   7935
         Begin MSComCtl2.DTPicker DTPFecFin 
            Height          =   252
            Left            =   4920
            TabIndex        =   51
            Top             =   720
            Visible         =   0   'False
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   445
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   37739
         End
         Begin MSComCtl2.DTPicker DTPFecIni 
            Height          =   252
            Left            =   3360
            TabIndex        =   50
            Top             =   720
            Visible         =   0   'False
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   445
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   37739
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Orden"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   41
            Top             =   1440
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Linea"
            Height          =   255
            Index           =   1
            Left            =   360
            Picture         =   "DetalleProduccionLiberada.frx":8AB3
            TabIndex        =   40
            Top             =   1080
            Width           =   1452
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            TabIndex        =   21
            Top             =   1920
            Width           =   2655
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Datos"
            Height          =   735
            Index           =   0
            Left            =   4920
            Picture         =   "DetalleProduccionLiberada.frx":B92D
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2400
            Width           =   2655
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Todos Los Datos"
            Height          =   735
            Index           =   1
            Left            =   4920
            Picture         =   "DetalleProduccionLiberada.frx":BEB7
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   3360
            Width           =   2655
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas"
            Height          =   255
            Index           =   0
            Left            =   360
            Picture         =   "DetalleProduccionLiberada.frx":C1C1
            TabIndex        =   20
            Top             =   720
            Width           =   855
         End
         Begin VB.Label LblFecFin 
            Caption         =   "Fecha Final"
            Height          =   252
            Left            =   4920
            TabIndex        =   53
            Top             =   480
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label LblFecIni 
            Caption         =   "Fecha Inicial"
            Height          =   252
            Left            =   3360
            TabIndex        =   52
            Top             =   480
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label LblDesPar 
            Alignment       =   1  'Right Justify
            Caption         =   "Orden"
            Height          =   255
            Left            =   3360
            TabIndex        =   26
            Top             =   1920
            Width           =   1455
         End
      End
      Begin VB.Frame FrameParos 
         Caption         =   "Datos del Produccion Liberada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3852
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   8175
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataParos"
            Height          =   285
            Index           =   6
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3120
            Width           =   6372
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Pasada"
            DataSource      =   "DataParos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   6
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FichaTecnica"
            DataSource      =   "DataParos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   5
            Top             =   1680
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Fecha"
            DataSource      =   "DataParos"
            Height          =   288
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   240
            Width           =   1692
            _ExtentX        =   2985
            _ExtentY        =   508
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataParos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            MaxLength       =   20
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Orden"
            DataSource      =   "DataParos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataParos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   2
            Top             =   600
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Turno"
            DataSource      =   "DataParos"
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   3
            Top             =   960
            Width           =   1692
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "ProductoConforme"
            DataSource      =   "DataParos"
            Height          =   288
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            Top             =   2400
            Width           =   1692
            _ExtentX        =   2985
            _ExtentY        =   508
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Desperdicio"
            DataSource      =   "DataParos"
            Height          =   288
            Index           =   2
            Left            =   1680
            TabIndex        =   8
            Top             =   2760
            Width           =   1692
            _ExtentX        =   2985
            _ExtentY        =   508
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   192
            Index           =   7
            Left            =   240
            TabIndex        =   49
            Top             =   3120
            Width           =   1104
         End
         Begin VB.Label LblPasada 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3480
            TabIndex        =   48
            Top             =   2040
            Width           =   4572
         End
         Begin VB.Label LblFichaTecnica 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3480
            TabIndex        =   47
            Top             =   1680
            Width           =   4572
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Desperdicio"
            Height          =   192
            Index           =   6
            Left            =   240
            TabIndex        =   46
            Top             =   2760
            Width           =   888
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Producto Conforme"
            Height          =   192
            Index           =   5
            Left            =   240
            TabIndex        =   45
            Top             =   2400
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Pasada"
            Height          =   192
            Index           =   4
            Left            =   240
            TabIndex        =   44
            Top             =   2040
            Width           =   576
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   192
            Index           =   3
            Left            =   240
            TabIndex        =   43
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
            Height          =   192
            Index           =   2
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   444
         End
         Begin VB.Label LblLinea 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3480
            TabIndex        =   29
            Top             =   600
            Width           =   4572
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   192
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   192
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   3480
            Width           =   660
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   192
            Index           =   19
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   396
         End
      End
      Begin MSDBGrid.DBGrid DBGridParos 
         Bindings        =   "DetalleProduccionLiberada.frx":F03B
         Height          =   4455
         Left            =   -74880
         OleObjectBlob   =   "DetalleProduccionLiberada.frx":F053
         TabIndex        =   19
         Top             =   720
         Width           =   8175
      End
   End
   Begin VB.Data DataParos 
      Caption         =   "Ficha Tecnica De Empleados"
      Connect         =   "Access"
      DatabaseName    =   "C:\erick\Amapro Metalenvases\metalenvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleProduccionOrdenLiberada"
      Top             =   6240
      Width           =   8385
   End
End
Attribute VB_Name = "DetalleProduccionLiberada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BLinea As Boolean
Dim BFichaTecnica As Boolean
Dim BPasada As Boolean

Dim VUltimaOrden As String
Dim Vlinea As String
Dim VPasada As String
Dim VTotalProductoConforme As Long
Dim RBuscaSaldo As Recordset

Dim RBuscaLinea As Recordset
Dim RBuscaPasada As Recordset
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaMaximo As Recordset
Dim RBuscaFichaOrden As Recordset


Sub botones()
    If Bandera = True Then
         FrameParos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         DataParos.Visible = False
         DBGridParos.Visible = False
         FrameBusqueda.Visible = False
    Else
         FrameParos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataParos.Visible = True
         DBGridParos.Visible = True
         FrameBusqueda.Visible = True
    End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
    

                'AGREGAR
                If Index = 0 Then
                                    On Error Resume Next
                                    DataParos.Recordset.AddNew
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                            Bandera = True
                                            botones
                                            Txttexto.Item(3).Text = GUsuario
                                            Msk.Item(0).Text = Date
                                            Msk.Item(0).SetFocus
                                            
                                            'BUSCA EL CODIGO DE PARO MAXIMO
                                         '   Set RBuscaMaximo = Db.OpenRecordset("Select Max(Val(Codigo)) From DetalleProduccionOrdenLiberada")
                                         '       If RBuscaMaximo.RecordCount > 0 Then
                                         '           TxtTexto.Item(0).Text = RBuscaMaximo(0) + 1
                                         '       Else
                                         '           TxtTexto.Item(0).Text = ""
                                         '       End If
                                            
                                    End If
                                    
                'EDITAR
                ElseIf Index = 1 Then
                                    On Error Resume Next
                                    DataParos.Recordset.Edit
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                            Bandera = True
                                            botones
                                            Txttexto.Item(3).Text = GUsuario
                                            Msk.Item(0).SetFocus
                                    End If
                                    
                'GRABAR
                ElseIf Index = 2 Then
                                   On Error Resume Next
                                   'VERIFICA LA FECHA
                                   If Not IsDate(Msk.Item(0).Text) Then
                                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                                        Msk.Item(0).SetFocus
                                        Exit Sub
                                   End If
                                   
                                   Vlinea = Txttexto.Item(0).Text
                                   VUltimaOrden = Txttexto.Item(2).Text
                                   VPasada = Txttexto.Item(5).Text
                                   VTotalProductoConforme = Msk.Item(1).Text
                                   
                                   
                                    'BUSCA EL DETALLE DE LA ORDEN DE LA PRODUCCION Y BUSCA LO REQUERIDO
                                     Set RBuscaSaldo = Db.OpenRecordset("Select Entregado, Saldo From DetalleOrdenProduccion Where Documento = '" & VUltimaOrden & "' And Linea = '" & Vlinea & "' And Pasada = '" & VPasada & "'")
                                         If RBuscaSaldo.RecordCount > 0 Then
                                         Else
                                            MsgBox "Para Esta Linea y Orden y Pasada No se a ABIERTO una Orden De Produccion", vbOKOnly + vbInformation, "Verifique"
                                            Exit Sub
                                         End If
                                    
                                    'GRABA DATOS
                                    DataParos.Recordset.Update
                                                                        
                                    If Err <> 0 And Err <> 3022 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                       Err.Clear
                                    Else
                                                    'BUSCA EL DETALLE DE LA ORDEN DE LA PRODUCCION Y BUSCA LO REQUERIDO
                                                    Set RBuscaSaldo = Db.OpenRecordset("Select Entregado, Saldo From DetalleOrdenProduccion Where Documento = '" & VUltimaOrden & "' And Linea = '" & Vlinea & "' And Pasada = '" & VPasada & "'")
                                                    
                                                        If RBuscaSaldo.RecordCount > 0 Then
                                                            RBuscaSaldo.Edit
                                                                RBuscaSaldo!Entregado = RBuscaSaldo!Entregado + VTotalProductoConforme
                                                                RBuscaSaldo!Saldo = RBuscaSaldo!Saldo - VTotalProductoConforme
                                                            RBuscaSaldo.Update
                                                        End If
                        
                                    
                                       Bandera = False
                                       botones
                                       CmdBotones.Item(0).SetFocus
                                   End If
                'CANCELAR
                ElseIf Index = 3 Then
                                    On Error Resume Next
                                    DataParos.Recordset.CancelUpdate
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                        Bandera = False
                                        botones
                                    End If
                'BORRAR
                ElseIf Index = 4 Then
                                    On Error Resume Next
                                            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                                
                                            If mensaje = vbOK Then
                                            
                                                Vlinea = Txttexto.Item(0).Text
                                                VUltimaOrden = Txttexto.Item(2).Text
                                                VPasada = Txttexto.Item(5).Text
                                                VTotalProductoConforme = Msk.Item(1).Text
                                                
                                                'BORRA EL REGISTRO
                                                DataParos.Recordset.Delete
                                                
                                                If Err <> 0 Then
                                                   MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                Else
                                                'BUSCA EL DETALLE DE LA ORDEN DE LA PRODUCCION Y BUSCA LO REQUERIDO
                                                    Set RBuscaSaldo = Db.OpenRecordset("Select Entregado, Saldo From DetalleOrdenProduccion Where Documento = '" & VUltimaOrden & "' And Linea = '" & Vlinea & "' And Pasada = '" & VPasada & "'")
                                                    
                                                        If RBuscaSaldo.RecordCount > 0 Then
                                                            RBuscaSaldo.Edit
                                                                RBuscaSaldo!Entregado = RBuscaSaldo!Entregado - VTotalProductoConforme
                                                                RBuscaSaldo!Saldo = RBuscaSaldo!Saldo + VTotalProductoConforme
                                                            RBuscaSaldo.Update
                                                        End If
                        
                                                
                                                End If
                                                DataParos.Recordset.MoveLast
                                            End If
                                  
                                            If DataParos.Recordset.EOF Then
                                                DataParos.Recordset.MoveLast
                                                If Err = 3021 Then
                                                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                                                End If
                                            End If
                                           
                                           If Err <> 0 Then
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                           End If
                                            
                'SALIDA
                Else
                                        Unload Me
                End If
    
    
End Sub

Private Sub CmdBusqueda_Click(Index As Integer)
MousePointer = 11
        If Index = 0 Then
            If OptOpcion.Item(0).Value = True Then
                DataParos.RecordSource = "Select * from DetalleProduccionOrdenLiberada Where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "#"
            ElseIf OptOpcion.Item(1).Value = True Then
                DataParos.RecordSource = "Select * from DetalleProduccionOrdenLiberada Where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "# And Linea = '" & TxtBusqueda.Text & "'"
            ElseIf OptOpcion.Item(1).Value = True Then
                DataParos.RecordSource = "Select * from DetalleProduccionOrdenLiberada Where Orden = '" & TxtBusqueda.Text & "'"
            End If
            DataParos.Refresh
            DBGridParos.Refresh
        End If
        
        If Index = 1 Then
            DataParos.RecordSource = "Select * from DetalleProduccionOrdenLiberada"
            DataParos.Refresh
            DBGridParos.Refresh
        End If
            TabParos.Tab = 1

MousePointer = 0
End Sub

Private Sub CmdSale_Click()
    Framebuscar.Visible = False
End Sub

Private Sub DBGridBuscar_DblClick()
    If BLinea = True Then
        Txttexto.Item(0).Text = DBGridBuscar.Columns(0)
        Txttexto.Item(0).SetFocus
    ElseIf BFichaTecnica = True Then
        Txttexto.Item(4).Text = DBGridBuscar.Columns(0)
        Txttexto.Item(4).SetFocus
    ElseIf BPasada = True Then
        Txttexto.Item(5).Text = DBGridBuscar.Columns(0)
        Txttexto.Item(5).SetFocus
    End If
        Framebuscar.Visible = False

End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BLinea = True Then
                    Txttexto.Item(0).Text = DBGridBuscar.Columns(0)
                    Txttexto.Item(0).SetFocus
                ElseIf BFichaTecnica = True Then
                    Txttexto.Item(4).Text = DBGridBuscar.Columns(0)
                    Txttexto.Item(4).SetFocus
                ElseIf BPasada = True Then
                    Txttexto.Item(5).Text = DBGridBuscar.Columns(0)
                    Txttexto.Item(5).SetFocus
                End If
                    Framebuscar.Visible = False
        End If
End Sub
Private Sub dbgridparos_HeadClick(ByVal ColIndex As Integer)
    DataParos.RecordSource = ("Select * from DetalleProduccionOrdenLiberada order by " & DBGridParos.Columns(ColIndex).DataField)
    DataParos.Refresh
    DBGridParos.Refresh
End Sub
Private Sub Form_Load()
    DataParos.Connect = GConnect
    DataParos.DatabaseName = BasedeDatos
    DataBuscar.Connect = GConnect
    DataBuscar.DatabaseName = BasedeDatos
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


Private Sub OptOpcion_Click(Index As Integer)
    If OptOpcion.Item(0).Value = True Then
        LblDesPar.Caption = ""
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        DtpFecIni.Visible = True
        DtpFecFin.Visible = True
    ElseIf OptOpcion.Item(1).Value = True Then
        LblDesPar.Caption = "Linea"
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        DtpFecIni.Visible = True
        DtpFecFin.Visible = True
        TxtBusqueda.SetFocus
    ElseIf OptOpcion.Item(2).Value = True Then
        LblDesPar.Caption = "Orden"
        LblFecIni.Visible = False
        LblFecFin.Visible = False
        DtpFecIni.Visible = False
        DtpFecFin.Visible = False
        TxtBusqueda.SetFocus
    End If
    
End Sub

Private Sub TabParos_Click(PreviousTab As Integer)
    If TabParos.Tab = 2 Then
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date
        TxtBusqueda.SetFocus
    End If
End Sub

Private Sub TxtBuscar_Change()
            
            
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                            'OPCION CUALQUIER PALABRA
                            If OptBusqueda.Item(3).Value = True Then
                                    If BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip From Lineas Where Descrip Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BFichaTecnica = True Then
                                        DataBuscar.RecordSource = ("Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BPasada = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion From Pasadas Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                                    End If
                            'OPCION PALABRA INICIAL
                            ElseIf OptBusqueda.Item(2).Value = True Then
                                    If BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip From Lineas Where Descrip Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BFichaTecnica = True Then
                                        DataBuscar.RecordSource = ("Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BPasada = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion From Pasadas Where Descripcion Like '" & TxtBuscar.Text & "*'")
                                    End If
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION CUALQUIER PALABRA
                            If OptBusqueda.Item(3).Value = True Then
                                    If BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip From Lineas Where Linea Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BFichaTecnica = True Then
                                        DataBuscar.RecordSource = ("Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BPasada = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion From Pasadas Where Codigo Like '*" & TxtBuscar.Text & "*'")
                                    End If
                            'OPCION PALABRA INICIAL
                            ElseIf OptBusqueda.Item(2).Value = True Then
                                    If BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip From Lineas Where Linea Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BFichaTecnica = True Then
                                        DataBuscar.RecordSource = ("Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BPasada = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion From Pasadas Where Descripcion Like '" & TxtBuscar.Text & "*'")
                                    End If
                            End If
                    End If
                            DataBuscar.Refresh
                            DBGridBuscar.Refresh
                            DBGridBuscar.Columns(1).Width = "5000"
                            
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



Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
    'LINEA
    If Index = 0 Then
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txttexto.Item(0).Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
    'FICHA TECNICA
    ElseIf Index = 4 Then
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & Txttexto.Item(4).Text & "'")
            If RBuscaFichaTecnica.RecordCount > 0 Then
                LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
            Else
                LblFichaTecnica.Caption = ""
            End If
    'PASADA
    ElseIf Index = 5 Then
        Set RBuscaPasada = Db.OpenRecordset("Select Descripcion From Pasadas Where Codigo = '" & Txttexto.Item(5).Text & "'")
            If RBuscaPasada.RecordCount > 0 Then
                LblPasada.Caption = RBuscaPasada!Descripcion
            Else
                LblPasada.Caption = ""
            End If
            
    End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        'LINEAS
        If Index = 0 Then
            BLinea = True
            BFichaTecnica = False
            BPasada = False
            DataBuscar.RecordSource = "Select Linea, Descrip from Lineas"
        'FICHA TECNICA
        ElseIf Index = 4 Then
            BLinea = False
            BFichaTecnica = True
            BPasada = False
            DataBuscar.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1"
        'PASADAS
        ElseIf Index = 5 Then
            BLinea = False
            BFichaTecnica = False
            BPasada = True
            DataBuscar.RecordSource = "Select Codigo, Descripcion From Pasadas"
        End If
        
        If Index = 0 Or Index = 4 Or Index = 5 Then
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DBGridBuscar.Columns(1).Width = "4000"
        End If

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        Txttexto.Item(Index).SelStart = 0
        Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        If Index = 0 Then
            BLinea = True
            BFichaTecnica = False
            BPasada = False
            DataBuscar.RecordSource = "Select Linea, Descrip from Lineas"
        ElseIf Index = 4 Then
            BLinea = False
            BFichaTecnica = True
            BPasada = False
            DataBuscar.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1"
        ElseIf Index = 5 Then
            BLinea = False
            BFichaTecnica = False
            BPasada = True
            DataBuscar.RecordSource = "Select Codigo, Descripcion From Pasadas"
        End If
        
        If Index = 0 Or Index = 4 Or Index = 5 Then
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DBGridBuscar.Columns(1).Width = "4000"
        End If
    End If
    
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
    If Index = 2 Then
        Set RBuscaFichaOrden = Db.OpenRecordset("Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & Txttexto.Item(2).Text & "'")
                    If RBuscaFichaOrden.RecordCount > 0 Then
                        Txttexto.Item(4).Text = RBuscaFichaOrden!FichaTecnica
                    Else
                        Txttexto.Item(4).Text = ""
                    End If
    End If
End Sub
