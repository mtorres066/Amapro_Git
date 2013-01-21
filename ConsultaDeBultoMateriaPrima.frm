VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form ConsultaDeBultoMateriaPrima 
   BackColor       =   &H80000004&
   Caption         =   "Consulta Especial De Bulto De Materia Prima"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ConsultaDeBultoMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataAjustes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8175
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Data DataConsultas 
         Caption         =   "consultas"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   5040
         TabIndex        =   28
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   8640
         Picture         =   "ConsultaDeBultoMateriaPrima.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "ConsultaDeBultoMateriaPrima.frx":293C
         Height          =   6975
         Left            =   120
         OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":2958
         TabIndex        =   27
         ToolTipText     =   "Signo '+' o Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   11535
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxtCodMatPri 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Signo '+' o Doble Click Para Seleccionar"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtTexto 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   5
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Data DataConsumos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDBGrid.DBGrid DBGridConsumos 
      Bindings        =   "ConsultaDeBultoMateriaPrima.frx":3333
      Height          =   1815
      Left            =   2760
      OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":334E
      TabIndex        =   40
      Top             =   4680
      Width           =   5895
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   975
   End
   Begin VB.Data DataCerrarBulto 
      Caption         =   "Cerrar Bulto"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data DataTraslados 
      Caption         =   "Traslados"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGridCerrarBulto 
      Bindings        =   "ConsultaDeBultoMateriaPrima.frx":3D48
      Height          =   1800
      Left            =   120
      OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":3D66
      TabIndex        =   34
      Top             =   6480
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGridTraslados 
      Bindings        =   "ConsultaDeBultoMateriaPrima.frx":4763
      Height          =   1920
      Left            =   120
      OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":477F
      TabIndex        =   33
      Top             =   840
      Width           =   11655
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
      Height          =   372
      Left            =   11400
      Picture         =   "ConsultaDeBultoMateriaPrima.frx":5176
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "salida"
      Top             =   0
      Width           =   372
   End
   Begin VB.CommandButton CmdGenerar 
      Height          =   372
      Left            =   10920
      Picture         =   "ConsultaDeBultoMateriaPrima.frx":71E8
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "consultar"
      Top             =   0
      Width           =   372
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   255
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   11
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox TxtTexto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox TxtNumIng 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Signo '+' o Doble Click Para Seleccionar"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Data DataDespachos 
      Caption         =   "Despachos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGridDespachos 
      Bindings        =   "ConsultaDeBultoMateriaPrima.frx":925A
      Height          =   1920
      Left            =   2760
      OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":9276
      TabIndex        =   36
      Top             =   2760
      Width           =   9015
   End
   Begin MSDBGrid.DBGrid DBGridAjustes 
      Bindings        =   "ConsultaDeBultoMateriaPrima.frx":9C69
      Height          =   1815
      Left            =   8640
      OleObjectBlob   =   "ConsultaDeBultoMateriaPrima.frx":9C83
      TabIndex        =   48
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Peso En Kilos"
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
      Left            =   240
      TabIndex        =   47
      Top             =   6000
      Width           =   1185
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Transaccion"
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
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
      Index           =   8
      Left            =   8040
      TabIndex        =   45
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label LblObs 
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
      Left            =   9360
      TabIndex        =   44
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ORDEN"
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
      Left            =   8640
      TabIndex        =   41
      Top             =   120
      Width           =   660
   End
   Begin VB.Label LblDocumento 
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
      Height          =   615
      Left            =   240
      TabIndex        =   39
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "No. Documento"
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
      Left            =   240
      TabIndex        =   38
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label LblUniMedPes 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   240
      TabIndex        =   35
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Entradas De Materia Prima"
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
      TabIndex        =   32
      Top             =   2760
      Width           =   2520
   End
   Begin VB.Label LblBodDis 
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
      Left            =   4920
      TabIndex        =   21
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LblCodMatPri 
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
      Left            =   3600
      TabIndex        =   20
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
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
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   17
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Bodega Actual"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Calidad"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Cantidad Entrada"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   1230
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Fecha De Entrada"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1305
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Numero De Bulto"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label Lbletiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Codigo Materia Prima"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   3375
      Index           =   0
      Left            =   120
      Top             =   3000
      Width           =   2535
   End
End
Attribute VB_Name = "ConsultaDeBultoMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaEntradas As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaBodegaEntrada As Recordset
Dim RBuscaBodegaDisponible As Recordset
Dim RBuscaDespachos As Recordset

Dim Columnas As String
Dim Tablas As String
Dim Criteria As String
Dim Columnas2 As String
Dim Tablas2 As String
Dim Criteria2 As String
Dim Columnas3 As String
Dim Tablas3 As String
Dim Criteria3 As String
Dim Columnas4 As String
Dim Tablas4 As String
Dim Criteria4 As String
Dim Columnas5 As String
Dim Tablas5 As String
Dim Criteria5 As String

Dim Columnas6 As String
Dim Tablas6 As String
Dim Criteria6 As String

Dim Columnas7 As String
Dim Tablas7 As String
Dim Criteria7 As String


'RECORDSET PARA BUSCAR DATOS DE TRASLADOS
Dim RTraslados As Recordset
Dim RDevoluciones As Recordset
Dim RDespachos As Recordset

'RECORDSET PARA BUSCAR DATOS DE CIERRE DE BULTO
Dim RCerrarBulto As Recordset

Dim BMateriaPrima As Boolean
Dim BNumeroIngreso As Boolean


Private Sub CmdGenerar_Click()
On Error Resume Next

        If Not IsNumeric(TxtNumIng.Text) Then
            MsgBox "Numero De Bulto Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
            TxtNumIng.SetFocus
            Exit Sub
        End If
                
        'BUSCA LAS ENTRADAS
        Columnas = "EE.FechaEntrada, EE.Documento, DE.Cantidad, DE.PesoEntrada, DE.Calidad, DE.BodegaDisponibilidad, DE.CantidadTraslado, DE.CantidadSalida, DE.SaldoDisponibilidad, DE.Peso, EE.NumeroDocumento, D.Descripcion, DE.OrdenProduccion, DE.Observaciones"
        Tablas = "EncabezadoEntradasMateriaPrima as EE, DetalleEntradasMateriaPrima as DE, Documentos as D"
        Criteria = "DE.Codigo = '" & TxtCodMatPri.Text & "' And DE.NumeroIngreso = " & TxtNumIng.Text & " And EE.Documento = DE.Documento And EE.TipoDeDocumento = D.CodigoDocumento"
        
        Set RBuscaEntradas = Db.OpenRecordset("Select " & Columnas & " From " & Tablas & " Where " & Criteria)
            
            If RBuscaEntradas.RecordCount > 0 Then
                    Txttexto.Item(0).Text = RBuscaEntradas(0)
                    Txttexto.Item(1).Text = RBuscaEntradas(1)
                    Txttexto.Item(2).Text = RBuscaEntradas(10)
                    Txttexto.Item(3).Text = RBuscaEntradas(2)
                    Txttexto.Item(4).Text = RBuscaEntradas(3)
                    Txttexto.Item(6).Text = RBuscaEntradas(4)
                    Txttexto.Item(8).Text = RBuscaEntradas(5)
                    Txttexto.Item(9).Text = RBuscaEntradas(6)
                    Txttexto.Item(10).Text = RBuscaEntradas(7)
                    Txttexto.Item(11).Text = RBuscaEntradas(8)
                    Txttexto.Item(13).Text = RBuscaEntradas(9)
                    LblDocumento.Caption = RBuscaEntradas(11)
                    Txttexto.Item(5).Text = RBuscaEntradas(12)
                    LblObs.Caption = RBuscaEntradas(13)
                    
'BUSCA LOS TRASLADOS DEL BULTO______________________________________________________________________
                    Columnas2 = "ET.Fecha, ET.NumeroDocumento, D.Descripcion, ET.Estado, ET.BodegaSalida, DT.CantidadSalida, B.Descripcion, DT.DiferenciaReqCorMas, DT.DiferenciaReqCor, DT.CantidadDesperdicio, DT.CantidadDesperdicioProveedor, DT.CantidadReal"
                    Tablas2 = "EncabezadoTrasladosMateriaPrim As ET, DetalleTrasladosMateriaPrimaP As DT, Documentos as D, BodegasMateriaPrima as B"
                    Criteria2 = "DT.CodigoSalida = '" & TxtCodMatPri.Text & "' And DT.NumeroIngreso = " & TxtNumIng.Text & " And DT.Documento = ET.Documento And ET.TipoDeDocumento = D.CodigoDocumento And DT.BodegaEntrada = B.CodigoBodega Order BY ET.Fecha"
                    DataTraslados.RecordSource = "Select " & Columnas2 & " From " & Tablas2 & " Where " & Criteria2
                    DataTraslados.Refresh
                    DBGridTraslados.Refresh
                    ColumnasTraslados
                        
'BUSCA LOS CIERRES DE BULTO_________________________________________________________________________
                    Columnas3 = "N.Fecha, N.Hora, L.Descrip, B.Descripcion, N.Existencia, N.CantidadProcesada, N.DesperdicioProceso, N.DesperdicioProveedor, N.CantidadProcesadaReal, N.CantidadMas, N.CantidadMenos, N.Total, N.UsuarioAgregar"
                    Tablas3 = "NumerosIngresosProcesados as N, Lineas as L, BodegasMateriaPrima as B"
                    Criteria3 = "N.CodigoMateriaPrima = '" & TxtCodMatPri.Text & "' And N.NumeroIngreso = " & TxtNumIng.Text & " And N.Linea = L.Linea And N.BodegaSalida = B.CodigoBodega Order By N.Fecha"
                    DataCerrarBulto.RecordSource = "Select " & Columnas3 & " From " & Tablas3 & " Where " & Criteria3
                    DataCerrarBulto.Refresh
                    DBGridCerrarBulto.Refresh
                    ColumnasCerrarBulto
                                        
                            
'BUSCA LOS DESPACHOS DEL BULTO_________________________________________________________________________
                    Columnas4 = "EMP.NumeroDocumento, D.Descripcion, EMP.Fecha, C.Descripcion, DMP.Bodega, DMP.Cantidad, EMP.Observaciones"
                    Tablas4 = "EncabezadoEgresosMateriaPrima as EMP, DetalleEgresosMateriaPrima as DMP, Clientes as C, Documentos as D"
                    Criteria4 = "EMP.Documento = DMP.Documento And EMP.Cliente = C.CodigoCliente And DMP.Codigo = '" & TxtCodMatPri.Text & "' And NumeroIngreso = " & TxtNumIng.Text & " And EMP.TipoDeDocumento = D.CodigoDocumento Order By EMP.Fecha"
                    DataDespachos.RecordSource = "Select " & Columnas4 & " From " & Tablas4 & " Where " & Criteria4
                    DataDespachos.Refresh
                    DBGridDespachos.Refresh
                    ColumnasDespachos
                            
            Else
                    Txttexto.Item(0).Text = ""
                    Txttexto.Item(1).Text = ""
                    Txttexto.Item(3).Text = ""
                    Txttexto.Item(4).Text = ""
                    Txttexto.Item(6).Text = ""
                    Txttexto.Item(8).Text = ""
                    Txttexto.Item(9).Text = ""
                    Txttexto.Item(10).Text = ""
                    Txttexto.Item(11).Text = ""
                    
                    DBGridTraslados.ClearFields
                    DBGridCerrarBulto.ClearFields
                    
                    MsgBox "Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
            End If
            
'CONSUMO DE MATERIAS PRIMAS EN REPORTE DE PRODUCCION
'BUSCA LOS CIERRES DE BULTO_________________________________________________________________________
                    Columnas6 = "DC.Documento, DC.Orden, DC.Desperdicio, DC.Cantidad, C.UnidadMedida"
                    Tablas6 = "DetalleConsumoMateriaPrima as DC, CorrelativosMateriaPrima as C"
                    Criteria6 = "DC.CodigoMateriaPrima = '" & TxtCodMatPri.Text & "' And DC.NumeroIngreso = " & TxtNumIng.Text & " And DC.CodigoMateriaPrima = C.CodigoMateriaPrima"
                    DataConsumos.RecordSource = "Select " & Columnas6 & " From " & Tablas6 & " Where " & Criteria6
                    DataConsumos.Refresh
                    DBGridConsumos.Refresh
                    ColumnasConsumos
                    
'CONSUMO DE MATERIAS PRIMAS EN REPORTE DE PRODUCCION
'BUSCA LOS AJUSTES _________________________________________________________________________
                    Columnas7 = "Fecha, Efecto, Cantidad, Observaciones, Usuario"
                    Tablas7 = "AjustesMateriaPrima"
                    Criteria7 = "CodigoMateriaPrima = '" & TxtCodMatPri.Text & "' And NumeroIngreso = " & TxtNumIng.Text
                    DataAjustes.RecordSource = "Select " & Columnas7 & " From " & Tablas7 & " Where " & Criteria7
                    DataAjustes.Refresh
                    DBGridAjustes.Refresh
                    ColumnasAjustes
                    

        
End Sub

Private Sub CmdSale_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub


Private Sub DBGridConsultas_DblClick()
        'MATERIAS PRIMAS
        If BMateriaPrima = True Then
            TxtCodMatPri.Text = DBGridConsultas.Columns(0).Text
            TxtCodMatPri.SetFocus
        'NUMERO DE INGRESO
        ElseIf BNumeroIngreso = True Then
            TxtNumIng.Text = DBGridConsultas.Columns(1).Text
            TxtNumIng.SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                'MATERIAS PRIMAS
                If BMateriaPrima = True Then
                    TxtCodMatPri.Text = DBGridConsultas.Columns(0).Text
                    TxtCodMatPri.SetFocus
                'NUMERO DE INGRESO
                ElseIf BNumeroIngreso = True Then
                    TxtNumIng.Text = DBGridConsultas.Columns(1).Text
                    TxtNumIng.SetFocus
                End If
                FrameConsultas.Visible = False
        End If
End Sub

Private Sub Form_Load()
        DataConsultas.ConnectionString = GTipoProveedor
        DataTraslados.ConnectionString = GTipoProveedor
        DataCerrarBulto.ConnectionString = GTipoProveedor
        DataDespachos.ConnectionString = GTipoProveedor
        DataConsumos.ConnectionString = GTipoProveedor
        DataAjustes.ConnectionString = GTipoProveedor

        DataConsultas.Refresh
        DataDespachos.Refresh
        DataTraslados.Refresh
        DataCerrarBulto.Refresh
        DataConsumos.Refresh
        DataAjustes.Refresh
End Sub

Private Sub TxtCodMatPri_Change()
        'BUSCA LA DESCRIPCION DEL CODIGO DE MATERIA PRIMA
        Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedidaPeso From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodMatPri.Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblCodMatPri.Caption = RBuscaMateriaPrima!Descripcion
                LblUniMedPes.Caption = RBuscaMateriaPrima!UnidadMedidaPeso
            Else
                LblCodMatPri.Caption = ""
                LblUniMedPes.Caption = ""
            End If
End Sub

Private Sub TxtCodMatPri_DblClick()
            
            BMateriaPrima = True
            BNumeroIngreso = False
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion from CorrelativosMateriaPrima"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
            DBGridConsultas.Columns(1).Width = "4000"

End Sub

Private Sub TxtCodMatPri_GotFocus()
        TxtCodMatPri.SelStart = 0
        TxtCodMatPri.SelLength = Len(TxtCodMatPri.Text)
End Sub

Private Sub TxtCodMatPri_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BMateriaPrima = True
            BNumeroIngreso = False
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion from CorrelativosMateriaPrima"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
            DBGridConsultas.Columns(1).Width = "4000"
        End If
End Sub

Private Sub TxtConsultas_Change()
        'MATERIA PRIMA
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima"
            End If
        End If
    
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    

End Sub

Private Sub TxtConsultas_GotFocus()
        TxtConsultas.SelStart = 0
        TxtConsultas.SelLength = Len(TxtConsultas.Text)
End Sub

Private Sub TxtConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtNumIng_DblClick()
        BMateriaPrima = False
        BNumeroIngreso = True
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataConsultas.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, Cantidad, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodMatPri.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        FrameConsultas.Visible = True
        TxtConsultas.SetFocus
        

End Sub

Private Sub TxtNumIng_GotFocus()
        TxtNumIng.SelStart = 0
        TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                BMateriaPrima = False
                BNumeroIngreso = True
                'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
                DataConsultas.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, Cantidad, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodMatPri.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                FrameConsultas.Visible = True
                TxtConsultas.SetFocus
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 8 Then
            Set RBuscaBodegaDisponible = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & Txttexto.Item(8).Text & "'")
                If RBuscaBodegaDisponible.RecordCount > 0 Then
                    LblBodDis.Caption = RBuscaBodegaDisponible!Descripcion
                Else
                    LblBodDis.Caption = ""
                End If
        End If
            
End Sub

Sub ColumnasTraslados()
                    DBGridTraslados.Columns(0).Width = "1000"
                    DBGridTraslados.Columns(1).Width = "800"
                    DBGridTraslados.Columns(2).Width = "2200"
                    DBGridTraslados.Columns(3).Width = "700"
                    DBGridTraslados.Columns(4).Width = "700"
                    DBGridTraslados.Columns(5).Width = "900"
                    DBGridTraslados.Columns(6).Width = "1900"
                    DBGridTraslados.Columns(7).Width = "500"
                    DBGridTraslados.Columns(8).Width = "500"
                    DBGridTraslados.Columns(9).Width = "500"
                    DBGridTraslados.Columns(10).Width = "500"
                    DBGridTraslados.Columns(11).Width = "900"
                    
                    
                    DBGridTraslados.Columns(0).Caption = "Fecha"
                    DBGridTraslados.Columns(1).Caption = "# Documento"
                    DBGridTraslados.Columns(2).Caption = "Documento"
                    DBGridTraslados.Columns(3).Caption = "Estado"
                    DBGridTraslados.Columns(4).Caption = "Bod.Sal."
                    DBGridTraslados.Columns(5).Caption = "Cant.Sal."
                    DBGridTraslados.Columns(6).Caption = "Bodega Entrada"
                    DBGridTraslados.Columns(7).Caption = "Cant. +"
                    DBGridTraslados.Columns(8).Caption = "Cant. -"
                    DBGridTraslados.Columns(9).Caption = "Des.Pro."
                    DBGridTraslados.Columns(10).Caption = "Des.Prov."
                    DBGridTraslados.Columns(11).Caption = "Cant.Real"
                    
                    DBGridTraslados.Columns(5).NumberFormat = "#,###,##0"
                    DBGridTraslados.Columns(7).NumberFormat = "#,###,##0"
                    DBGridTraslados.Columns(8).NumberFormat = "#,###,##0"
                    DBGridTraslados.Columns(9).NumberFormat = "#,###,##0"
                    DBGridTraslados.Columns(10).NumberFormat = "#,###,##0"
                    DBGridTraslados.Columns(11).NumberFormat = "#,###,##0"
                    
                    
End Sub

Sub ColumnasCerrarBulto()
                    DBGridCerrarBulto.Columns(0).Width = "1000"
                    DBGridCerrarBulto.Columns(1).Width = "600"
                    DBGridCerrarBulto.Columns(1).NumberFormat = "hh:mm"
                    DBGridCerrarBulto.Columns(2).Width = "1500"
                    DBGridCerrarBulto.Columns(3).Width = "1300"
                    DBGridCerrarBulto.Columns(4).Width = "800"
                    
                    DBGridCerrarBulto.Columns(5).Width = "800"
                    DBGridCerrarBulto.Columns(6).Width = "400"
                    DBGridCerrarBulto.Columns(7).Width = "400"
                    DBGridCerrarBulto.Columns(8).Width = "800"
                    DBGridCerrarBulto.Columns(9).Width = "800"
                    DBGridCerrarBulto.Columns(10).Width = "800"
                    
                    DBGridCerrarBulto.Columns(11).Width = "800"
                    DBGridCerrarBulto.Columns(12).Width = "1000"
                    
                    
                    DBGridCerrarBulto.Columns(0).Caption = "Fecha"
                    DBGridCerrarBulto.Columns(1).Caption = "Hora"
                    DBGridCerrarBulto.Columns(2).Caption = "Linea"
                    DBGridCerrarBulto.Columns(3).Caption = "Bodega Salida"
                    DBGridCerrarBulto.Columns(4).Caption = "Requizado"
                    
                    DBGridCerrarBulto.Columns(5).Caption = "Cant.Proc."
                    DBGridCerrarBulto.Columns(6).Caption = "Des.Proc."
                    DBGridCerrarBulto.Columns(7).Caption = "Des.Prov."
                    DBGridCerrarBulto.Columns(8).Caption = "Cant.Proc.Real"
                    DBGridCerrarBulto.Columns(9).Caption = "Cant. +"
                    DBGridCerrarBulto.Columns(10).Caption = "Cant. -"
                    
                    DBGridCerrarBulto.Columns(11).Caption = "Descargar"
                    DBGridCerrarBulto.Columns(12).Caption = "Usuario"
                    
End Sub


Sub ColumnasDespachos()
                    DBGridDespachos.Columns(0).Width = "800"
                    DBGridDespachos.Columns(1).Width = "1700"
                    DBGridDespachos.Columns(2).Width = "1000"
                    DBGridDespachos.Columns(3).Width = "1500"
                    DBGridDespachos.Columns(4).Width = "500"
                    DBGridDespachos.Columns(5).Width = "1000"
                    DBGridDespachos.Columns(6).Width = "2000"
                    
                    DBGridDespachos.Columns(0).Caption = "# Documento"
                    DBGridDespachos.Columns(1).Caption = "Documento"
                    DBGridDespachos.Columns(2).Caption = "Fecha"
                    DBGridDespachos.Columns(3).Caption = "Cliente"
                    DBGridDespachos.Columns(4).Caption = "Bodega"
                    DBGridDespachos.Columns(5).Caption = "Cantidad"
                    DBGridDespachos.Columns(5).NumberFormat = "#,###,##0.00"
                    DBGridDespachos.Columns(6).Caption = "Observaciones"
End Sub

Public Sub ColumnasAjustes()
                    DBGridAjustes.Columns(0).Width = "900"
                    DBGridAjustes.Columns(1).Width = "200"
                    DBGridAjustes.Columns(2).Width = "800"
                    DBGridAjustes.Columns(3).Width = "2500"
                    DBGridAjustes.Columns(4).Width = "1000"
                    
                    DBGridAjustes.Columns(0).Caption = "Fecha"
                    DBGridAjustes.Columns(0).NumberFormat = "dd/mm/yy"
                    DBGridAjustes.Columns(1).Caption = "Efecto"
                    DBGridAjustes.Columns(2).Caption = "Cantidad"
                    DBGridAjustes.Columns(2).NumberFormat = "#,###,##0.00"
                    DBGridAjustes.Columns(3).Caption = "Observaciones"
                    DBGridAjustes.Columns(4).Caption = "Usuario"
                    
End Sub

Public Sub ColumnasConsumos()
                    DBGridConsumos.Columns(0).Width = "1000"
                    DBGridConsumos.Columns(1).Width = "1500"
                    DBGridConsumos.Columns(2).Width = "1000"
                    DBGridConsumos.Columns(3).Width = "1000"
                    DBGridConsumos.Columns(4).Width = "1000"
                    
                    DBGridConsumos.Columns(0).Caption = "Documento"
                    DBGridConsumos.Columns(1).Caption = "Orden"
                    DBGridConsumos.Columns(2).Caption = "Desperdicio"
                    DBGridConsumos.Columns(3).Caption = "Cantidad"
                    DBGridConsumos.Columns(4).Caption = "U/Medida"
End Sub
