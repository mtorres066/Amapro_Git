VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ConsultaDeProduccion 
   BackColor       =   &H000080FF&
   Caption         =   "Consulta De Reporte De Produccion"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "ConsultaDeProduccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Data DataBusqueda 
         Caption         =   "Busqueda"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "ConsultaDeProduccion.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "ConsultaDeProduccion.frx":21BC
         Height          =   6975
         Left            =   120
         OleObjectBlob   =   "ConsultaDeProduccion.frx":21D7
         TabIndex        =   21
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin VB.Data DataInvProTer 
      Caption         =   "Inv Pro Ter"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Data DataInvMatPri 
      Caption         =   "Inv. Mat. Pri."
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.Data Dataparos 
      Caption         =   "Paros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Data DataLineas 
      Caption         =   "Lineas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Data DataGerencia 
      Caption         =   "Gerencia"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Tipo De Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1812
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H000080FF&
         Caption         =   "Todos"
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
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptLinea 
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptGrupo 
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
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
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11280
      Picture         =   "ConsultaDeProduccion.frx":2BB1
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   0
      Width           =   495
   End
   Begin VB.Data DataOrden 
      Caption         =   "Orden"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Data DataMes 
      Caption         =   "Mes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1212
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   252
      Left            =   8040
      TabIndex        =   1
      Top             =   120
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61603843
      CurrentDate     =   37248
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   252
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61603843
      CurrentDate     =   37248
   End
   Begin TabDlg.SSTab TabGeneral 
      Height          =   7452
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   11892
      _ExtentX        =   20981
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Produccion"
      TabPicture(0)   =   "ConsultaDeProduccion.frx":4C23
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGridMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DBGridParos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DBGridLineas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DBGridGerencia"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ordenes Abiertas"
      TabPicture(1)   =   "ConsultaDeProduccion.frx":4F3D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabProduccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab TabProduccion 
         Height          =   6612
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   11652
         _ExtentX        =   20558
         _ExtentY        =   11668
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   706
         BackColor       =   8438015
         ForeColor       =   16711680
         TabCaption(0)   =   "Resumen Ordenes Abiertas"
         TabPicture(0)   =   "ConsultaDeProduccion.frx":6C47
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FGrid"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detalle Ordenes Abiertas"
         TabPicture(1)   =   "ConsultaDeProduccion.frx":6C63
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DBGridOrden"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Inventario Materia Prima"
         TabPicture(2)   =   "ConsultaDeProduccion.frx":6C7F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "DBGridInvMatPri"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Inventario Producto Terminado"
         TabPicture(3)   =   "ConsultaDeProduccion.frx":6C9B
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "DBGridInvProTer"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin MSDBGrid.DBGrid DBGridInvProTer 
            Bindings        =   "ConsultaDeProduccion.frx":6CB7
            Height          =   5895
            Left            =   -74880
            OleObjectBlob   =   "ConsultaDeProduccion.frx":6CD3
            TabIndex        =   24
            Top             =   600
            Width           =   11415
         End
         Begin MSDBGrid.DBGrid DBGridInvMatPri 
            Bindings        =   "ConsultaDeProduccion.frx":76AE
            Height          =   5895
            Left            =   -74880
            OleObjectBlob   =   "ConsultaDeProduccion.frx":76CA
            TabIndex        =   25
            Top             =   600
            Width           =   11415
         End
         Begin MSFlexGridLib.MSFlexGrid FGrid 
            Height          =   5895
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Doble Click Para Ver Detalle"
            Top             =   600
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10398
            _Version        =   393216
            Rows            =   10
            Cols            =   9
            WordWrap        =   -1  'True
            FocusRect       =   2
            HighLight       =   2
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin MSDBGrid.DBGrid DBGridOrden 
            Bindings        =   "ConsultaDeProduccion.frx":80A5
            Height          =   5895
            Left            =   -74880
            OleObjectBlob   =   "ConsultaDeProduccion.frx":80BD
            TabIndex        =   27
            Top             =   600
            Width           =   11415
         End
      End
      Begin MSDBGrid.DBGrid DBGridGerencia 
         Bindings        =   "ConsultaDeProduccion.frx":8A94
         Height          =   3732
         Left            =   120
         OleObjectBlob   =   "ConsultaDeProduccion.frx":8AAF
         TabIndex        =   28
         Tag             =   "Produccion Agrupada Por Ficha Tecnica"
         Top             =   720
         Width           =   7752
      End
      Begin MSDBGrid.DBGrid DBGridLineas 
         Bindings        =   "ConsultaDeProduccion.frx":949D
         Height          =   2772
         Left            =   120
         OleObjectBlob   =   "ConsultaDeProduccion.frx":94B6
         TabIndex        =   29
         Tag             =   "Produccion Agrupada Por Linea"
         Top             =   4560
         Width           =   5412
      End
      Begin MSDBGrid.DBGrid DBGridParos 
         Bindings        =   "ConsultaDeProduccion.frx":9EA2
         Height          =   3732
         Left            =   7920
         OleObjectBlob   =   "ConsultaDeProduccion.frx":9EBA
         TabIndex        =   30
         Top             =   720
         Width           =   3852
      End
      Begin MSDBGrid.DBGrid DBGridMes 
         Bindings        =   "ConsultaDeProduccion.frx":A89D
         Height          =   2772
         Left            =   5640
         OleObjectBlob   =   "ConsultaDeProduccion.frx":A8B3
         TabIndex        =   31
         Tag             =   "Produccion Agrupada Por Mes"
         Top             =   4560
         Width           =   6132
      End
   End
   Begin VB.Label LblLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label LblDescripcion 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
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
      Left            =   5640
      TabIndex        =   7
      Top             =   480
      Width           =   552
   End
   Begin MSForms.CommandButton CmdGenera 
      Default         =   -1  'True
      Height          =   495
      Left            =   10680
      TabIndex        =   4
      ToolTipText     =   "Generar Datos"
      Top             =   0
      Width           =   495
      BackColor       =   12632256
      PicturePosition =   327683
      Size            =   "873;873"
      Picture         =   "ConsultaDeProduccion.frx":B298
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Al"
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
      Index           =   1
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Del"
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
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "ConsultaDeProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As Recordset
Dim RBuscaLinea As Recordset
Dim ROrdenesResumen As Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean
Dim Cont As Integer
Dim VTotalFilas As Integer
Dim VOrden As String



Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11

'_______________________________________________________________________________________________________________________
            'GRID DE FICHA TECNICA
            
                If OptTodos.Value = True Then
                    DataGerencia.RecordSource = "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos as EP, DetalleProduccionPorOrden As P, FichaTecnica As F, EncabezadoOrdenProduccion as EO Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataGerencia.RecordSource = "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos as EP, DetalleProduccionPorOrden As P, FichaTecnica As F, EncabezadoOrdenProduccion as EO, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataGerencia.RecordSource = "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos as EP, DetalleProduccionPorOrden As P, FichaTecnica As F, EncabezadoOrdenProduccion as EO, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip"
                End If
            DataGerencia.Refresh
            DBGridGerencia.Refresh
            DBGridGerencia.Columns(2).NumberFormat = "#,###,##0"
            DBGridGerencia.Columns(3).NumberFormat = "#,###,##0"
            DBGridGerencia.Columns(4).NumberFormat = "#,###,##0"
            DBGridGerencia.Columns(0).Caption = "Orden"
            DBGridGerencia.Columns(1).Caption = "Ficha Tecnica"
            DBGridGerencia.Columns(2).Caption = "PC"
            DBGridGerencia.Columns(3).Caption = "PNC"
            DBGridGerencia.Columns(4).Caption = "Desp."
            DBGridGerencia.Columns(0).Width = "1300"
            DBGridGerencia.Columns(1).Width = "3500"
            DBGridGerencia.Columns(2).Width = "800"
            DBGridGerencia.Columns(3).Width = "800"
            DBGridGerencia.Columns(4).Width = "800"

'_______________________________________________________________________________________________________________________
            'EL GRID DE LINEAS
            
                If OptTodos.Value = True Then
                    DataLineas.RecordSource = "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea Group By EP.Linea, L.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataLineas.RecordSource = "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By EP.Linea, L.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataLineas.RecordSource = "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By EP.Linea, L.Descrip"
                End If
            
            DataLineas.Refresh
            DBGridLineas.Refresh
            DBGridLineas.Columns(2).NumberFormat = "#,###,##0"
            DBGridLineas.Columns(3).NumberFormat = "#,###,##0"
            DBGridLineas.Columns(4).NumberFormat = "#,###,##0"
            DBGridLineas.Columns(0).Caption = "Linea"
            DBGridLineas.Columns(1).Caption = "Descripcion"
            DBGridLineas.Columns(2).Caption = "PC"
            DBGridLineas.Columns(3).Caption = "PNC"
            DBGridLineas.Columns(4).Caption = "Desp."
            DBGridLineas.Columns(0).Width = "300"
            DBGridLineas.Columns(1).Width = "2100"
            DBGridLineas.Columns(2).Width = "800"
            DBGridLineas.Columns(3).Width = "800"
            DBGridLineas.Columns(4).Width = "800"

            
'_______________________________________________________________________________________________________________________
            'EL GRID DE MES
            
            
                If OptTodos.Value = True Then
                    DataMes.RecordSource = "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)"
                    'DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptGrupo.Value = True Then
                    DataMes.RecordSource = "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)"
                ElseIf OptLinea.Value = True Then
                    DataMes.RecordSource = "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos as EP, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)"
                End If
            
            DataMes.Refresh
            DBGridMes.Refresh
            DBGridMes.Columns(4).NumberFormat = "#,###,##0"
            DBGridMes.Columns(5).NumberFormat = "#,###,##0"
            DBGridMes.Columns(6).NumberFormat = "#,###,##0"
            DBGridMes.Columns(0).Caption = "Mes"
            DBGridMes.Columns(1).Caption = "Año"
            DBGridMes.Columns(2).Caption = "Linea"
            DBGridMes.Columns(3).Caption = "Descripcion"
            DBGridMes.Columns(4).Caption = "PC"
            DBGridMes.Columns(5).Caption = "PNC"
            DBGridMes.Columns(6).Caption = "Desp."
            DBGridMes.Columns(0).Width = "300"
            DBGridMes.Columns(1).Width = "400"
            DBGridMes.Columns(2).Width = "300"
            DBGridMes.Columns(3).Width = "2200"
            DBGridMes.Columns(4).Width = "800"
            DBGridMes.Columns(5).Width = "800"
            DBGridMes.Columns(6).Width = "800"


            
    '_______________________________________________________________________________________________________________________
            'PAROS
            
                If OptTodos.Value = True Then
                    DataParos.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
                ElseIf OptGrupo.Value = True Then
                    DataParos.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
                ElseIf OptLinea.Value = True Then
                    DataParos.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & TxtLinea.Text & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
                End If
            DataParos.Refresh
            DBGridParos.Refresh
            
            DBGridParos.Columns(3).NumberFormat = "#,###,##0.00"
            
            DBGridParos.Columns(0).Caption = "Linea"
            DBGridParos.Columns(1).Caption = "Descripcion"
            DBGridParos.Columns(2).Caption = "Tipo"
            DBGridParos.Columns(3).Caption = "Horas"
            
            DBGridParos.Columns(0).Width = "300"
            DBGridParos.Columns(1).Width = "2100"
            DBGridParos.Columns(2).Width = "200"
            DBGridParos.Columns(3).Width = "600"
            
            
            If Err <> 0 Then
                'MsgBox Err.Description
            End If
            
MousePointer = 0
        
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False

End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
            If BLinea = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(0)
            ElseIf BGrupo = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(2)
            End If
            FrameBusqueda.Visible = False
            TxtLinea.SetFocus
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BLinea = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(0)
                ElseIf BGrupo = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(2)
                End If
                FrameBusqueda.Visible = False
                TxtLinea.SetFocus
            End If
End Sub

Private Sub FGrid_DblClick()
On Error Resume Next

            'ASIGNAMOS A UNA VARIABLE LA ORDEN DE LA COLUMNA 1
            FGrid.Col = 1
            VOrden = FGrid.Text
            
            DBGridOrden.Caption = "ORDEN " & Space(5) & VOrden
            DBGridInvMatPri.Caption = "ORDEN " & Space(5) & VOrden
            DBGridInvProTer.Caption = Space(50) & "ORDEN: " & Space(10) & VOrden
                
            'EL GRID DE ORDEN DETALLE ABIERTAS
                        DataOrden.RecordSource = "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo From DetalleOrdenProduccion as DO, Lineas as L, Pasadas as P Where DO.Documento = '" & VOrden & "' And DO.Linea = L.Linea And DO.Pasada = P.Codigo"
                        DataOrden.Refresh
                        DBGridOrden.Refresh
                        
                        DBGridOrden.Columns(3).NumberFormat = "#,###,##0"
                        DBGridOrden.Columns(4).NumberFormat = "#,###,##0"
                        DBGridOrden.Columns(5).NumberFormat = "#,###,##0"
                        
                        DBGridOrden.Columns(0).Caption = "Linea"
                        DBGridOrden.Columns(1).Caption = "Pasada"
                        DBGridOrden.Columns(2).Caption = "Observacion"
                        DBGridOrden.Columns(3).Caption = "Requerido"
                        DBGridOrden.Columns(4).Caption = "Producido"
                        DBGridOrden.Columns(5).Caption = "Saldo"
                                  
                        
                        DBGridOrden.Columns(0).Width = "3000"
                        DBGridOrden.Columns(1).Width = "1500"
                        DBGridOrden.Columns(2).Width = "1500"
                        DBGridOrden.Columns(3).Width = "1000"
                        DBGridOrden.Columns(4).Width = "1000"
                        DBGridOrden.Columns(5).Width = "1000"
                        
                        TabProduccion.Tab = 1
            
           'INVENTARIO MATERIA PRIMA
                        DataInvMatPri.RecordSource = "Select DE.BodegaDisponibilidad, B.Descripcion, DE.Codigo, C.Descripcion, Count(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad / C.CuerposPorLamina), Sum(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad * C.PesoxUnidad / 1000) From DetalleEntradasMateriaPrima as DE, BodegasMateriaPrima as B, CorrelativosMateriaPrima as C Where DE.OrdenProduccion = '" & VOrden & "' And DE.SaldoDisponibilidad > 0 And DE.BodegaDisponibilidad = B.CodigoBodega And DE.Codigo = C.CodigoMateriaPrima Group By DE.BodegaDisponibilidad, B.Descripcion, DE.Codigo, C.Descripcion"
                        DataInvMatPri.Refresh
                        DBGridInvMatPri.Refresh
           
           
                        DBGridInvMatPri.Columns(4).NumberFormat = "#,###,##0"
                        DBGridInvMatPri.Columns(5).NumberFormat = "#,###,##0.00"
                        DBGridInvMatPri.Columns(6).NumberFormat = "#,###,##0"
                        DBGridInvMatPri.Columns(7).NumberFormat = "#,###,##0.00"
                        
                        DBGridInvMatPri.Columns(0).Caption = "Bodega"
                        DBGridInvMatPri.Columns(1).Caption = "Descripcion"
                        DBGridInvMatPri.Columns(2).Caption = "Materia Prima"
                        DBGridInvMatPri.Columns(3).Caption = "Descripcion"
                        DBGridInvMatPri.Columns(4).Caption = "Bultos"
                        DBGridInvMatPri.Columns(5).Caption = "Laminas"
                        DBGridInvMatPri.Columns(6).Caption = "Unidades"
                        DBGridInvMatPri.Columns(7).Caption = "Toneladas"
                                   
                        DBGridInvMatPri.Columns(0).Width = "500"
                        DBGridInvMatPri.Columns(1).Width = "2500"
                        DBGridInvMatPri.Columns(2).Width = "1500"
                        DBGridInvMatPri.Columns(3).Width = "2500"
                        DBGridInvMatPri.Columns(4).Width = "1000"
                        DBGridInvMatPri.Columns(5).Width = "1000"
                        DBGridInvMatPri.Columns(6).Width = "1000"
                        DBGridInvMatPri.Columns(7).Width = "1000"
                        
                        
           'INVENTARIO PRODUCTO TERMINADO
                        DataInvProTer.RecordSource = "Select DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip, Count(DE.Saldo), Sum(DE.Saldo / F.UnidadesxLamina), Sum(DE.Saldo), Sum(DE.Saldo * F.PesoxUnidad / 1000) From DetalleEntradasProductoTermina as DE, BodegasProductoTerminado as B, FichaTecnica as F Where DE.OrdenProduccion = '" & VOrden & "' And DE.Saldo > 0 And DE.Bodega = B.CodigoBodega And DE.FichaTecnica = F.Esp_Tec Group By DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip"
                        DataInvProTer.Refresh
                        DBGridInvProTer.Refresh
           
           
                        DBGridInvProTer.Columns(4).NumberFormat = "#,###,##0"
                        DBGridInvProTer.Columns(5).NumberFormat = "#,###,##0.00"
                        DBGridInvProTer.Columns(6).NumberFormat = "#,###,##0"
                        DBGridInvProTer.Columns(7).NumberFormat = "#,###,##0.00"
                        
                        DBGridInvProTer.Columns(0).Caption = "Bodega"
                        DBGridInvProTer.Columns(1).Caption = "Descripcion"
                        DBGridInvProTer.Columns(2).Caption = "Materia Prima"
                        DBGridInvProTer.Columns(3).Caption = "Descripcion"
                        DBGridInvProTer.Columns(4).Caption = "Bultos"
                        DBGridInvProTer.Columns(5).Caption = "Laminas"
                        DBGridInvProTer.Columns(6).Caption = "Unidades"
                        DBGridInvProTer.Columns(7).Caption = "Toneladas"
                                   
                        DBGridInvProTer.Columns(0).Width = "500"
                        DBGridInvProTer.Columns(1).Width = "2500"
                        DBGridInvProTer.Columns(2).Width = "1500"
                        DBGridInvProTer.Columns(3).Width = "2500"
                        DBGridInvProTer.Columns(4).Width = "1000"
                        DBGridInvProTer.Columns(5).Width = "1000"
                        DBGridInvProTer.Columns(6).Width = "1000"
                        DBGridInvProTer.Columns(7).Width = "1000"
           
End Sub

Private Sub Form_Load()
            DataGerencia.ConnectionString = GTipoProveedor
            DataLineas.ConnectionString = GTipoProveedor
            DataMes.ConnectionString = GTipoProveedor
            DataOrden.ConnectionString = GTipoProveedor
            DataBusqueda.ConnectionString = GTipoProveedor
            DataParos.ConnectionString = GTipoProveedor
            DataInvMatPri.ConnectionString = GTipoProveedor
            DataInvProTer.ConnectionString = GTipoProveedor

            DataGerencia.Refresh
            DataLineas.Refresh
            DataMes.Refresh
            DataParos.Refresh
            DataOrden.Refresh
            DataBusqueda.Refresh
            DataInvMatPri.Refresh
            DataInvProTer.Refresh
            
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
            CmdGenera_Click
End Sub

Private Sub Form_Resize()
            'DBGridGerencia.Height = Me.ScaleHeight - 5100
'            DBGridMes.Height = Me.ScaleHeight - 5500
'            DBGridOrden.Height = Me.ScaleHeight - 4700
            
End Sub

Private Sub OptGrupo_Click()
            LblDescripcion.Caption = "Grupo"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub

Private Sub OptLinea_Click()
            LblDescripcion.Caption = "Linea"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub

Private Sub OptTodos_Click()
            LblDescripcion.Caption = ""
            TxtLinea.Visible = False
End Sub

Private Sub TabGeneral_Click(PreviousTab As Integer)
    
        If TabGeneral.Tab = 1 Then
        
            Cont = 1
            '_______________________________________________________________________________________________________________________
            'EL GRID DE ORDEN RESUMEN
            
            'Set ROrdenesResumen = Db.OpenRecordset("Select EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega, Sum(DO.Requerido), Sum(DO.Entregado), Sum(DO.Saldo) From EncabezadoOrdenProduccion as EO, DetalleOrdenProduccion as DO, FichaTecnica as F Where EO.Documento = DO.Documento And EO.FichaTecnica = F.Esp_Tec And EO.Estado = 'ABIERTA' Group By EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega Having Sum(DO.Saldo) > 0 Order by EO.Documento")
            Set ROrdenesResumen = Db.OpenRecordset("Select EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega, Sum(DO.Requerido), Sum(DO.Entregado), Sum(DO.Saldo) From EncabezadoOrdenProduccion as EO, DetalleOrdenProduccion as DO, FichaTecnica as F Where EO.Documento = DO.Documento And EO.FichaTecnica = F.Esp_Tec And EO.Estado = 'ABIERTA' Group By EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega Order by EO.Documento")
                
                If ROrdenesResumen.RecordCount > 0 Then
                                'SE MUEVE AL ULTIMO REGISTRO
                                ROrdenesResumen.MoveLast
                                'ASIGNA A UNA VARIABLE EL TOTAL DE REGISTROS
                                VTotalFilas = ROrdenesResumen.RecordCount + 1
                                'LIMPIA EL GRID
                                FGrid.Clear
                                'ASIGNA CUANTAS FILA VA A TENER EL FLEX GRID
                                FGrid.Rows = VTotalFilas
                                'REGRESA AL PRIMER REGISTRO PARA EMPEZAR A DESPLEGAR LOS DATOS
                                ROrdenesResumen.MoveFirst
                                
                                FGrid.Row = 0
                                FGrid.Col = 0
                                FGrid.ColWidth(0) = "100"
                                FGrid.Col = 1
                                FGrid.ColWidth(1) = "1500"
                                FGrid.Text = "Orden"
                                FGrid.Col = 2
                                FGrid.ColWidth(2) = "1500"
                                FGrid.Text = "Ficha Tecnica"
                                FGrid.Col = 3
                                FGrid.ColWidth(3) = "3000"
                                FGrid.Text = "Descripcion"
                                FGrid.Col = 4
                                FGrid.ColWidth(4) = "1000"
                                FGrid.Text = "Apertura"
                                FGrid.Col = 5
                                FGrid.ColWidth(5) = "1000"
                                FGrid.Text = "Entrega"
                                FGrid.Col = 6
                                FGrid.ColWidth(6) = "900"
                                FGrid.Text = "Requerido"
                                FGrid.Col = 7
                                FGrid.ColWidth(7) = "900"
                                FGrid.Text = "Producido"
                                FGrid.Col = 8
                                FGrid.ColWidth(8) = "900"
                                FGrid.Text = "Saldo"
                                                
                    
                        Do Until ROrdenesResumen.EOF
                                FGrid.Row = Cont
                                FGrid.Col = 1
                                FGrid.CellBackColor = vbGreen
                                FGrid.Text = ROrdenesResumen(0)
                                FGrid.Col = 2
                                FGrid.Text = ROrdenesResumen(1)
                                FGrid.Col = 3
                                FGrid.Text = ROrdenesResumen(2)
                                FGrid.Col = 4
                                FGrid.Text = ROrdenesResumen(3)
                                FGrid.Col = 5
                                FGrid.Text = ROrdenesResumen(4)
                                FGrid.Col = 6
                                FGrid.CellBackColor = vbYellow
                                FGrid.Text = Format(ROrdenesResumen(5), "#,###,##0")
                                FGrid.Col = 7
                                FGrid.Text = Format(ROrdenesResumen(6), "#,###,##0")
                                FGrid.Col = 8
                                FGrid.CellBackColor = vbCyan
                                FGrid.Text = Format(ROrdenesResumen(7), "#,###,##0")
                                
                                Cont = Cont + 1
                            ROrdenesResumen.MoveNext
                        Loop
                End If
        End If

End Sub

Private Sub Txtbusqueda_Change()
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Descrip Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Descrip Like '*" & Txtbusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Linea Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Linea Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtLinea_Change()
        If OptLinea.Value = True Then
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        End If
            
End Sub

Private Sub TxtLinea_DblClick()
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
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
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
End Sub
