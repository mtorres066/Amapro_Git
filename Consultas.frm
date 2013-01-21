VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Consultas 
   Caption         =   "Consultas Generales"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "Consultas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data DataConsultas 
      Caption         =   "Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGridConsultas 
      Bindings        =   "Consultas.frx":0442
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "Consultas.frx":045E
      TabIndex        =   3
      Top             =   2520
      Width           =   11775
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
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
      Height          =   975
      Left            =   10680
      Picture         =   "Consultas.frx":0E39
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "&Consultar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      Picture         =   "Consultas.frx":1143
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab TabConsultas 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4207
      _Version        =   393216
      Tabs            =   16
      TabsPerRow      =   7
      TabHeight       =   882
      TabCaption(0)   =   "Ficha Tecnica Envases"
      TabPicture(0)   =   "Consultas.frx":1585
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "OptEnvCod"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "OptEnvPla"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OptEnvFor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OptEnvFon"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptEnvTap"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptEnvDes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Ficha Tecnica Fondos"
      TabPicture(1)   =   "Consultas.frx":15A1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptFonBar"
      Tab(1).Control(1)=   "OptFonDes"
      Tab(1).Control(2)=   "OptFonCod"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ficha Tecnica Tapas"
      TabPicture(2)   =   "Consultas.frx":15BD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OptTapBar"
      Tab(2).Control(1)=   "OptTapFor"
      Tab(2).Control(2)=   "OptTapDes"
      Tab(2).Control(3)=   "OptTapCod"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Ficha Tecnica Platinas"
      TabPicture(3)   =   "Consultas.frx":15D9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "OptPlaBar"
      Tab(3).Control(1)=   "OptPlaDes"
      Tab(3).Control(2)=   "OptPlaCod"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Ficha Tecnica Rutinas"
      TabPicture(4)   =   "Consultas.frx":15F5
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "OptRutDes"
      Tab(4).Control(1)=   "OptRutCod"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tipos de Formas"
      TabPicture(5)   =   "Consultas.frx":1611
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "OptForDes"
      Tab(5).Control(1)=   "OptForCod"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Tipos de Barniz"
      TabPicture(6)   =   "Consultas.frx":162D
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "OptBarDes"
      Tab(6).Control(1)=   "OptBarCod"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Tipos de Alambre"
      TabPicture(7)   =   "Consultas.frx":1649
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "OptAlaDes"
      Tab(7).Control(1)=   "OptAlaCod"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Tipos de Barniz Liquido"
      TabPicture(8)   =   "Consultas.frx":1665
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "OptBarLiqDes"
      Tab(8).Control(1)=   "OptBarLiqCod"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Tipos de Barniz Polvo"
      TabPicture(9)   =   "Consultas.frx":1681
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "OptBarPolDes"
      Tab(9).Control(1)=   "OptBarPolCod"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Tipos de Sello Solvente"
      TabPicture(10)  =   "Consultas.frx":169D
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "OptSelSolDes"
      Tab(10).Control(1)=   "OptselSolCod"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Tipos de Nylon Strech"
      TabPicture(11)  =   "Consultas.frx":16B9
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "OptNylStrDes"
      Tab(11).Control(1)=   "OptNylStrCod"
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "Tipos de Defectos"
      TabPicture(12)  =   "Consultas.frx":16D5
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "OptDefTip"
      Tab(12).Control(1)=   "OptDefDes"
      Tab(12).Control(2)=   "OptDefCod"
      Tab(12).ControlCount=   3
      TabCaption(13)  =   "Variables"
      TabPicture(13)  =   "Consultas.frx":16F1
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "OptVarCod"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "Produccion"
      TabPicture(14)  =   "Consultas.frx":170D
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "PProFecFin"
      Tab(14).Control(1)=   "PProFecIni"
      Tab(14).Control(2)=   "OptProBat"
      Tab(14).Control(3)=   "OptProFicTec"
      Tab(14).Control(4)=   "OptProFecHor"
      Tab(14).Control(5)=   "OptProFec"
      Tab(14).ControlCount=   6
      TabCaption(15)  =   "Rutinas"
      TabPicture(15)  =   "Consultas.frx":1729
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "PRutFecFin"
      Tab(15).Control(1)=   "PRutFecIni"
      Tab(15).Control(2)=   "OptRutFicTec"
      Tab(15).Control(3)=   "OptRutFecHor"
      Tab(15).Control(4)=   "OptRutFec"
      Tab(15).ControlCount=   5
      Begin VB.OptionButton OptVarCod 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   -74640
         TabIndex        =   51
         Top             =   1920
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker PRutFecFin 
         Height          =   255
         Left            =   -69240
         TabIndex        =   50
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   36907
      End
      Begin MSComCtl2.DTPicker PRutFecIni 
         Height          =   255
         Left            =   -70560
         TabIndex        =   49
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   36907
      End
      Begin VB.OptionButton OptRutFicTec 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Left            =   -72120
         TabIndex        =   48
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton OptRutFecHor 
         Caption         =   "Fecha Y Hora"
         Height          =   195
         Left            =   -73680
         TabIndex        =   47
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton OptRutFec 
         Caption         =   "Fechas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   46
         Top             =   1800
         Width           =   975
      End
      Begin MSComCtl2.DTPicker PProFecFin 
         Height          =   255
         Left            =   -68640
         TabIndex        =   45
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   36907
      End
      Begin MSComCtl2.DTPicker PProFecIni 
         Height          =   255
         Left            =   -70080
         TabIndex        =   44
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   36907
      End
      Begin VB.OptionButton OptProBat 
         Caption         =   "Batch"
         Height          =   195
         Left            =   -71040
         TabIndex        =   43
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton OptProFicTec 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Left            =   -72480
         TabIndex        =   42
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton OptProFecHor 
         Caption         =   "Fecha y hora"
         Height          =   195
         Left            =   -73800
         TabIndex        =   41
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton OptProFec 
         Caption         =   "Fechas"
         Height          =   195
         Left            =   -74760
         TabIndex        =   40
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton OptDefTip 
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -72360
         TabIndex        =   39
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptDefDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73800
         TabIndex        =   38
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptDefCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   37
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptNylStrDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73560
         TabIndex        =   36
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptNylStrCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   35
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptSelSolDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73680
         TabIndex        =   34
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptselSolCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   33
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptBarPolDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73680
         TabIndex        =   32
         Top             =   1815
         Width           =   1335
      End
      Begin VB.OptionButton OptBarPolCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptBarLiqDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73440
         TabIndex        =   30
         Top             =   1815
         Width           =   1575
      End
      Begin VB.OptionButton OptBarLiqCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74760
         TabIndex        =   29
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptAlaDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73560
         TabIndex        =   28
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptAlaCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   27
         Top             =   1815
         Width           =   855
      End
      Begin VB.OptionButton OptBarDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73680
         TabIndex        =   26
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptBarCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   25
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptForDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73440
         TabIndex        =   24
         Top             =   1815
         Width           =   1215
      End
      Begin VB.OptionButton OptForCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74760
         TabIndex        =   23
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptRutDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73680
         TabIndex        =   22
         Top             =   1815
         Width           =   1455
      End
      Begin VB.OptionButton OptRutCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   21
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptPlaBar 
         Caption         =   "Barniz"
         Height          =   195
         Left            =   -72600
         TabIndex        =   20
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptPlaDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73920
         TabIndex        =   19
         Top             =   1815
         Width           =   1335
      End
      Begin VB.OptionButton OptPlaCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptTapBar 
         Caption         =   "Barniz"
         Height          =   195
         Left            =   -71880
         TabIndex        =   17
         Top             =   1815
         Width           =   855
      End
      Begin VB.OptionButton OptTapFor 
         Caption         =   "Forma"
         Height          =   195
         Left            =   -72720
         TabIndex        =   16
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptTapDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73920
         TabIndex        =   15
         Top             =   1815
         Width           =   1335
      End
      Begin VB.OptionButton OptTapCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptFonBar 
         Caption         =   "Barniz"
         Height          =   195
         Left            =   -72000
         TabIndex        =   13
         Top             =   1815
         Width           =   1095
      End
      Begin VB.OptionButton OptFonDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   -73560
         TabIndex        =   12
         Top             =   1815
         Width           =   1335
      End
      Begin VB.OptionButton OptFonCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74760
         TabIndex        =   11
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptEnvDes 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   1815
         Width           =   1215
      End
      Begin VB.OptionButton OptEnvTap 
         Caption         =   "Tapa"
         Height          =   195
         Left            =   4800
         TabIndex        =   9
         Top             =   1815
         Width           =   855
      End
      Begin VB.OptionButton OptEnvFon 
         Caption         =   "Fondo"
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   1815
         Width           =   855
      End
      Begin VB.OptionButton OptEnvFor 
         Caption         =   "Forma"
         Height          =   195
         Left            =   3120
         TabIndex        =   7
         Top             =   1815
         Width           =   855
      End
      Begin VB.OptionButton OptEnvPla 
         Caption         =   "Platina"
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   1815
         Width           =   975
      End
      Begin VB.OptionButton OptEnvCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1815
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Columnas As String
Dim Tablas As String
Dim Criteria As String

Private Sub CmdConsultar_Click()
MousePointer = 11

'ENVASES
If TabConsultas.Tab = 0 Then
    Columnas = "*"
    Tablas = "FichaTecnica"
    If OptEnvCod.Value = True Then
        Criteria = "Esp_Tec like '" & TxtTexto.Text & "*'"
    ElseIf OptEnvDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    ElseIf OptEnvPla.Value = True Then
        Criteria = "Platina like '" & TxtTexto.Text & "*'"
    ElseIf OptEnvFor.Value = True Then
        Criteria = "Forma like '" & TxtTexto.Text & "*'"
    ElseIf OptEnvFon.Value = True Then
        Criteria = "Fondo like '" & TxtTexto.Text & "*'"
    ElseIf OptEnvTap.Value = True Then
        Criteria = "tapa like '" & TxtTexto.Text & "*'"
    End If
'FONDOS
ElseIf TabConsultas.Tab = 1 Then
    Columnas = "*"
    Tablas = "Fondos"
    If OptFonCod.Value = True Then
        Criteria = "Fondo like '" & TxtTexto.Text & "*'"
    ElseIf OptFonDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    ElseIf OptFonBar.Value = True Then
        Criteria = "Barniz like '" & TxtTexto.Text & "*'"
    End If
'TAPAS
ElseIf TabConsultas.Tab = 2 Then
    Columnas = "*"
    Tablas = "Tapas"
    If OptTapCod.Value = True Then
        Criteria = "Tapa like '" & TxtTexto.Text & "*'"
    ElseIf OptTapDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    ElseIf OptTapFor.Value = True Then
        Criteria = "Forma like '" & TxtTexto.Text & "*'"
    ElseIf OptTapBar.Value = True Then
        Criteria = "Barniz like '" & TxtTexto.Text & "*'"
    End If
'PLATINAS
ElseIf TabConsultas.Tab = 3 Then
    Columnas = "*"
    Tablas = "Platinas"
    If OptPlaCod.Value = True Then
        Criteria = "platina like '" & TxtTexto.Text & "*'"
    ElseIf OptPlaDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    ElseIf OptPlaBar.Value = True Then
        Criteria = "Barniz like '" & TxtTexto.Text & "*'"
    End If
'RUTINAS
ElseIf TabConsultas.Tab = 4 Then
    Columnas = "*"
    Tablas = "Rutinas"
    If OptRutCod.Value = True Then
        Criteria = "Rutina like '" & TxtTexto.Text & "*'"
    ElseIf OptRutDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    End If
'TIPOS DE FORMAS
ElseIf TabConsultas.Tab = 5 Then
    Columnas = "*"
    Tablas = "Formas"
    If OptForCod.Value = True Then
        Criteria = "Forma like '" & TxtTexto.Text & "*'"
    ElseIf OptForDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    End If
'TIPOS DE BARNIZ
ElseIf TabConsultas.Tab = 6 Then
    Columnas = "*"
    Tablas = "Barniz"
    If OptBarCod.Value = True Then
        Criteria = "Barniz like '" & TxtTexto.Text & "*'"
    ElseIf OptBarDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    End If
'TIPOS DE ALAMBRES
ElseIf TabConsultas.Tab = 7 Then
    Columnas = "*"
    Tablas = "Alambre"
    If OptAlaCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    ElseIf OptAlaDes.Value = True Then
        Criteria = "Descripcion like '" & TxtTexto.Text & "*'"
    End If
'BARNIZ LIQUIDO
ElseIf TabConsultas.Tab = 8 Then
    Columnas = "*"
    Tablas = "BarnizLiquido"
    If OptBarLiqCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    ElseIf OptBarLiqDes.Value = True Then
        Criteria = "Descripcion like '" & TxtTexto.Text & "*'"
    End If
'BARNIZ POLVO
ElseIf TabConsultas.Tab = 9 Then
    Columnas = "*"
    Tablas = "BarnizPolvo"
    If OptBarPolCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    ElseIf OptBarPolDes.Value = True Then
        Criteria = "Descripcion like '" & TxtTexto.Text & "*'"
    End If
'SOLVENTE
ElseIf TabConsultas.Tab = 10 Then
    Columnas = "*"
    Tablas = "SelloSolvente"
    If OptselSolCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    ElseIf OptSelSolDes.Value = True Then
        Criteria = "Descripcion like '" & TxtTexto.Text & "*'"
    End If

'NYLON STRECH
ElseIf TabConsultas.Tab = 11 Then
    Columnas = "*"
    Tablas = "NylonStrech"
    If OptNylStrCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    ElseIf OptNylStrDes.Value = True Then
        Criteria = "Descripcion like '" & TxtTexto.Text & "*'"
    End If
'DEFECTOS
ElseIf TabConsultas.Tab = 12 Then
    Columnas = "*"
    Tablas = "Defectos"
    If OptDefCod.Value = True Then
        Criteria = "Defecto like '" & TxtTexto.Text & "*'"
    ElseIf OptDefDes.Value = True Then
        Criteria = "Descrip like '" & TxtTexto.Text & "*'"
    ElseIf OptDefTip.Value = True Then
        Criteria = "Tipo like '" & TxtTexto.Text & "*'"
    End If
'VARIABLES
ElseIf TabConsultas.Tab = 13 Then
    Columnas = "*"
    Tablas = "VariablesMedia"
    If OptVarCod.Value = True Then
        Criteria = "Codigo like '" & TxtTexto.Text & "*'"
    End If
'PRODUCCION
ElseIf TabConsultas.Tab = 14 Then
    Columnas = "*"
    Tablas = "Produccion"
    If OptProFec.Value = True Then
        Criteria = "Fec_prd >= #" & Format(PProFecIni.Value, "mm/dd/yyyy") & "#" & " and Fec_prd <= #" & Format(PProFecFin.Value, "mm/dd/yyyy") & "#"
    ElseIf OptProFecHor.Value = True Then
        Criteria = "Fec_prd >= #" & Format(PProFecIni.Value, "mm/dd/yyyy") & "#" & " and Fec_prd <= #" & Format(PProFecFin.Value, "mm/dd/yyyy") & "# and Hor_Prd >= '" & TxtTexto.Text & "'"
    ElseIf OptProFicTec.Value = True Then
        Criteria = "Esp_tec like '" & TxtTexto.Text & "*'"
    ElseIf OptProBat.Value = True Then
        'SI ESTA VACIO
        If TxtTexto.Text = "" Then
            MsgBox "Ingrese Un Numero de Batch ", vbOKOnly + vbInformation, "Informacion"
            TxtTexto.SetFocus
            MousePointer = 0
            Exit Sub
        End If
        'SI NO ES NUMERO
        If Not IsNumeric(TxtTexto.Text) Then
            MsgBox "Numero de Batch Solo Puede Ser Numerico ", vbOKOnly + vbInformation, "Informacion"
            TxtTexto.SetFocus
            MousePointer = 0
            Exit Sub
        End If
        Criteria = "Batch = " & TxtTexto.Text
    End If

'RUTINAS
ElseIf TabConsultas.Tab = 15 Then
    Columnas = "*"
    Tablas = "CapturaRutinas"
    If OptRutFec.Value = True Then
        Criteria = "Fec_rut >= #" & Format(PRutFecIni.Value, "mm/dd/yyyy") & "#" & " and Fec_Rut <= #" & Format(PRutFecFin.Value, "mm/dd/yyyy") & "#"
    ElseIf OptRutFecHor.Value = True Then
        Criteria = "Fec_Rut >= #" & Format(PRutFecIni.Value, "mm/dd/yyyy") & "#" & " and Fec_Rut <= #" & Format(PRutFecFin.Value, "mm/dd/yyyy") & "# and Hor_Rut >= '" & TxtTexto.Text & "'"
    ElseIf OptRutFicTec.Value = True Then
        Criteria = "Esp_tec like '" & TxtTexto.Text & "*'"
    End If
End If

DataConsultas.RecordSource = "Select " & Columnas & " From " & Tablas & " Where " & Criteria
DataConsultas.Refresh
DBGridConsultas.Refresh

MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DataConsultas.Connect = GConnect
    DataConsultas.DatabaseName = BasedeDatos
End Sub

Private Sub OptAlaCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptAlaDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptBarCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptBarDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptBarLiqCod_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptBarLiqDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptBarPolCod_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptBarPolDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptDefCod_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptDefDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptDefTip_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptEnvCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptEnvDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptEnvFon_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptEnvFor_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptEnvPla_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptEnvTap_Click()
TxtTexto.SetFocus
End Sub


Private Sub OptFonBar_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptFonCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptFonDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptForCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptForDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptNylStrCod_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptNylStrDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptPlaBar_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptPlaCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptPlaDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptProBat_Click()
    TxtTexto.Visible = True
    TxtTexto.SetFocus
End Sub

Private Sub OptProFec_Click()
    PProFecIni.Visible = True
    PProFecIni.Visible = True
    PProFecIni.Value = Date
    PProFecFin.Value = Date
    TxtTexto.Visible = False
End Sub

Private Sub OptProFecHor_Click()
    PProFecIni.Visible = True
    PProFecIni.Visible = True
    PProFecIni.Value = Date
    PProFecFin.Value = Date
    TxtTexto.Visible = True
    TxtTexto.SetFocus

End Sub

Private Sub OptProFicTec_Click()
    TxtTexto.Visible = True
    TxtTexto.SetFocus
End Sub

Private Sub OptRutCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptRutDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptRutFec_Click()
    PRutFecIni.Visible = True
    PRutFecIni.Visible = True
    PRutFecIni.Value = Date
    PRutFecFin.Value = Date
    TxtTexto.Visible = False

End Sub

Private Sub OptRutFecHor_Click()
    PRutFecIni.Visible = True
    PRutFecIni.Visible = True
    PRutFecIni.Value = Date
    PRutFecFin.Value = Date
    TxtTexto.Visible = True
    TxtTexto.SetFocus

End Sub

Private Sub OptselSolCod_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptSelSolDes_Click()
TxtTexto.SetFocus
End Sub

Private Sub OptTapBar_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptTapCod_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptTapDes_Click()
    TxtTexto.SetFocus
End Sub

Private Sub OptTapFor_Click()
    TxtTexto.SetFocus
End Sub

Private Sub TabConsultas_Click(PreviousTab As Integer)
    'ENVASES
    If TabConsultas.Tab = 0 Then
        TxtTexto.Visible = True
        OptEnvCod.Value = True
    'FONDOS
    ElseIf TabConsultas.Tab = 1 Then
        TxtTexto.Visible = True
        OptFonCod.Value = True
    'TAPAS
    ElseIf TabConsultas.Tab = 2 Then
        TxtTexto.Visible = True
        OptTapCod.Value = True
    'PLATINAS
    ElseIf TabConsultas.Tab = 3 Then
        TxtTexto.Visible = True
        OptPlaCod.Value = True
    'RUTINAS
    ElseIf TabConsultas.Tab = 4 Then
        TxtTexto.Visible = True
        OptRutCod.Value = True
    'TIPOS DE FORMA
    ElseIf TabConsultas.Tab = 5 Then
        TxtTexto.Visible = True
        OptForCod.Value = True
    'TIPOS DE BARNIZ
    ElseIf TabConsultas.Tab = 6 Then
        TxtTexto.Visible = True
        OptBarCod.Value = True
    'ALAMBRE
    ElseIf TabConsultas.Tab = 7 Then
        TxtTexto.Visible = True
        OptAlaCod.Value = True
    'BARNIZ LIQUIDO
    ElseIf TabConsultas.Tab = 8 Then
        TxtTexto.Visible = True
        OptBarLiqCod.Value = True
    'BARNIZ POLVO
    ElseIf TabConsultas.Tab = 9 Then
        TxtTexto.Visible = True
        OptBarPolCod.Value = True
    'SELLO SOLVENTE
    ElseIf TabConsultas.Tab = 10 Then
        TxtTexto.Visible = True
        OptselSolCod.Value = True
    'NYLON STRECH
    ElseIf TabConsultas.Tab = 11 Then
        TxtTexto.Visible = True
        OptNylStrCod.Value = True
    'DEFECTOS
    ElseIf TabConsultas.Tab = 12 Then
        TxtTexto.Visible = True
        OptDefCod.Value = True
    'VARIABLES
    ElseIf TabConsultas.Tab = 13 Then
        OptVarCod.Value = True
        TxtTexto.Visible = True
    'PRODUCCION
    ElseIf TabConsultas.Tab = 14 Then
        TxtTexto.Visible = True
        OptProFec.Value = True
    'RUTINAS
    ElseIf TabConsultas.Tab = 15 Then
        TxtTexto.Visible = True
        OptRutFec.Value = True
    End If
    
    
    
    
    
End Sub

