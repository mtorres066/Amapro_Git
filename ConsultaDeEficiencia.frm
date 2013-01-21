VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConsultaDeEficiencia 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta De Eficiencia"
   ClientHeight    =   8592
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11880
   Icon            =   "ConsultaDeEficiencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Height          =   4455
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton CmdSalir 
         Height          =   495
         Left            =   10800
         Picture         =   "ConsultaDeEficiencia.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Escuintla\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "ConsultaDeEficiencia.frx":2004
         Height          =   4095
         Left            =   120
         OleObjectBlob   =   "ConsultaDeEficiencia.frx":201F
         TabIndex        =   16
         Top             =   240
         Width           =   10575
      End
   End
   Begin VB.Frame FrameTotal 
      Caption         =   "Total De Minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   10080
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox TxtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paros Tipo N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paros Tipo S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1110
      End
   End
   Begin VB.Data DataDetalle 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data DataEncabezado 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGridDetalle 
      Bindings        =   "ConsultaDeEficiencia.frx":29F9
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "ConsultaDeEficiencia.frx":2A13
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   9855
   End
   Begin MSDBGrid.DBGrid DBGridEncabezado 
      Bindings        =   "ConsultaDeEficiencia.frx":3404
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "ConsultaDeEficiencia.frx":3421
      TabIndex        =   20
      ToolTipText     =   "doble click o flechas abajo y arriba despliega detalle "
      Top             =   2640
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.Data DataConsultas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   1320
      Top             =   4440
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   10440
      Picture         =   "ConsultaDeEficiencia.frx":3E19
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   9000
      Picture         =   "ConsultaDeEficiencia.frx":5E8B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   4255
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Ficha De Paros"
      TabPicture(0)   =   "ConsultaDeEficiencia.frx":7B85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblParos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "OptPar(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtPar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OptPar(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptPar(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptPar(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Captura De Paros"
      TabPicture(1)   =   "ConsultaDeEficiencia.frx":7E9F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptCapPar(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DTPFecFinPar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DTPFecIniPar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtCapPar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "OptCapPar(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "OptCapPar(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LblLinea"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LblCapPar"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LblFecFin"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblfecini"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.OptionButton OptPar 
         Caption         =   "Grupo"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Tipo De Paro"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas Y Linea"
         Height          =   195
         Index           =   1
         Left            =   -73680
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPFecFinPar 
         Height          =   255
         Left            =   -67560
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   445
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   23658499
         CurrentDate     =   37123
      End
      Begin MSComCtl2.DTPicker DTPFecIniPar 
         Height          =   255
         Left            =   -69960
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   445
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   23658499
         CurrentDate     =   37123
      End
      Begin VB.TextBox TxtCapPar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67560
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Documento"
         Height          =   195
         Index           =   2
         Left            =   -72000
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtPar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Codigo Paro"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label LblLinea 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -72600
         TabIndex        =   33
         Top             =   2040
         Width           =   6255
      End
      Begin VB.Label LblCapPar 
         Alignment       =   2  'Center
         Caption         =   "Documento"
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
         Left            =   -68640
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label LblFecFin 
         Caption         =   "Fecha Final"
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
         Left            =   -68640
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblfecini 
         Caption         =   "Fecha Inicial"
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
         Left            =   -71280
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblParos 
         AutoSize        =   -1  'True
         Caption         =   "Codigo De Paro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   3
         Top             =   1920
         Width           =   1350
      End
   End
   Begin MSDBGrid.DBGrid DBGridConsultas 
      Bindings        =   "ConsultaDeEficiencia.frx":9BA9
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "ConsultaDeEficiencia.frx":9BC5
      TabIndex        =   19
      Top             =   2640
      Width           =   11655
   End
End
Attribute VB_Name = "ConsultaDeEficiencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VDia As String
Dim VDia2 As String
Dim VMes As String
Dim VMes2 As String
Dim VAño As String
Dim VAño2 As String

Dim Criteria As String

Dim RBuscaLinea As Recordset
'EFICIENCIA
Dim RTiempoProgramadoD As Recordset
Dim VTiempoProgramadoD As Double

Dim RTiempoProgramadoN As Recordset
Dim VTiempoProgramadoN As Double

'PAROS QUE NO AFECTAN LA PRODUCCION
Dim RBuscaParosNoAfectanD As Recordset
Dim VParosND As Double

Dim RBuscaParosNoAfectanN As Recordset
Dim VParosNN As Double

'PARO QUE SI AFECTAN LA PRODUCCION
Dim RBuscaParosSiAfectanD As Recordset
Dim VParosSD As Double

Dim RBuscaParosSiAfectanN As Recordset
Dim VParosSN As Double


'TIEMPO REAL DE PRODUCCION
Dim VTiempoRealProducidoD As Double
Dim VTiempoRealProducidoN As Double

Dim VHorasProducidasDN As Double
Dim VParosDN As Double

Dim VPCD As Double
Dim VPNCD As Double
Dim VPDD As Double

Dim VPCN As Double
Dim VPNCN As Double
Dim VPDN As Double

Dim RProduccion As Recordset

Dim VTotalProduccion As Double
Dim VTotalProduccionD As Double
Dim VTotalProduccionN As Double
Dim VVelocidadPromedio As Integer
Dim VVelocidadPromedioD As Integer
Dim VVelocidadPromedioN As Integer

'CALCULOS DE EFICIENCIA
Dim VVelocidadTeoricaDia As Integer
Dim VVelocidadTeoricaNoche As Integer
Dim VVelocidadRealDia As Integer
Dim VVelocidadRealNoche As Integer

Dim VVelocidadTeoricaLinea As Integer

Dim VFactor1D As Double
Dim VFactor2D As Double
Dim VFactor3D As Double
Dim VFactor4D As Double
Dim VFactor5D As Double

Dim VFactor1N As Double
Dim VFactor2N As Double
Dim VFactor3N As Double
Dim VFactor4N As Double
Dim VFactor5N As Double

Dim VFactor1DN As Double
Dim VFactor2DN As Double
Dim VFactor3DN As Double
Dim VFactor4DN As Double
Dim VFactor5DN As Double

Dim VFactor1TDN As Double
Dim VFactor2TDN As Double
Dim VFactor3TDN As Double
Dim VFactor4TDN As Double
Dim VFactor5TDN As Double


Dim RLineas As Recordset


Dim VEficienciaRealD As Double
Dim VEficienciaRealN As Double

Dim RReporteEficiencia As Recordset

Dim Vlinea As String
Dim VFechaInicial As Date
Dim VFechaFinal As Date

Dim RSeleccionaLineas As Recordset

Dim VPorcentajeLinea As Double
Dim VPorcentajeRechazo As Double
Dim VPorcentajeDesperdicio As Double

Dim cont As Integer
'Dim VFactorUno As Double
Dim RBuscaDescripcionLinea As Recordset


'VARIABLES PARA BUSQUEDA DE DATOS
Dim BGrupos As Boolean
Dim BParos As Boolean
Dim BEficiencia As Boolean

Dim RCapturaParos As Recordset
Dim RTotalS As Recordset
Dim RTotalN As Recordset
Dim RTotalP As Recordset


Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11


  'MATERIAS PRIMAS
  If TabReportes.Tab = 0 Then
                                Paros
  ElseIf TabReportes.Tab = 1 Then
                                CapturaParos
  End If
  
                If Err > 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                End If
                
  MousePointer = 0
End Sub

Private Sub CmdSalida_Click()
    Unload Me
    
End Sub


Private Sub CmdSalir_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BParos = True Then
                TxtPar.Text = DBGridBusqueda.Columns(0)
                TxtPar.SetFocus
        End If
                FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
        If BParos = True Then
                TxtPar.Text = DBGridBusqueda.Columns(0)
                TxtPar.SetFocus
        End If
                FrameBusqueda.Visible = False
End If

End Sub

Private Sub DBGridEncabezado_DblClick()
            DataDetalle.RecordSource = "Select DC.Orden, DC.Inicio, DC.Final, DC.Minutos, DC.Paro, P.DescripcionParo From Paros as P, DetalleCapturaParos as DC Where DC.Documento = " & DBGridEncabezado.Columns(0) & " And DC.Paro = P.CodigoParo Order by DC.Inicio"
            DataDetalle.Refresh
            DBGridDetalle.Refresh
            ColumnasDetalle
            
            'SUMA LOS MINUTOS TIPO S
            Set RTotalS = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'S' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalN = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'N' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalP = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'P' And DC.Documento = " & DBGridEncabezado.Columns(0))
            
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
                    
End Sub

Private Sub DBGridEncabezado_KeyDown(KeyCode As Integer, Shift As Integer)
            DataDetalle.RecordSource = "Select DC.Orden, DC.FichaTecnica, DC.Inicio, DC.Final, DC.Minutos, DC.Paro, P.DescripcionParo From Paros as P, DetalleCapturaParos as DC Where DC.Documento = " & DBGridEncabezado.Columns(0) & " And DC.Paro = P.CodigoParo Order by DC.Inicio"
            DataDetalle.Refresh
            DBGridDetalle.Refresh
            ColumnasDetalle
            
            'SUMA LOS MINUTOS TIPO S
            Set RTotalS = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'S' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalN = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'N' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalP = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'P' And DC.Documento = " & DBGridEncabezado.Columns(0))
            
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
            
            
End Sub

Private Sub DBGridEncabezado_KeyUp(KeyCode As Integer, Shift As Integer)
            DataDetalle.RecordSource = "Select DC.Orden, DC.FichaTecnica, DC.Inicio, DC.Final, DC.Minutos, DC.Paro, P.DescripcionParo From Paros as P, DetalleCapturaParos as DC Where DC.Documento = " & DBGridEncabezado.Columns(0) & " And DC.Paro = P.CodigoParo Order by DC.Inicio"
            DataDetalle.Refresh
            DBGridDetalle.Refresh
            ColumnasDetalle
            
            'SUMA LOS MINUTOS TIPO S
            Set RTotalS = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'S' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalN = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'N' And DC.Documento = " & DBGridEncabezado.Columns(0))
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
            Set RTotalP = Db.OpenRecordset("Select sum(DC.minutos) from DetalleCapturaParos as DC, Paros As P where DC.Paro = P.CodigoParo And P.Tipo = 'P' And DC.Documento = " & DBGridEncabezado.Columns(0))
            
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
            
            
End Sub

Private Sub Form_Load()
            DataBusqueda.Connect = GConnect
            DataConsultas.Connect = GConnect
            DataEncabezado.Connect = GConnect
            DataDetalle.Connect = GConnect
            
            DataBusqueda.DatabaseName = BasedeDatos
            DataConsultas.DatabaseName = BasedeDatos
            DataEncabezado.DatabaseName = BasedeDatos
            DataDetalle.DatabaseName = BasedeDatos
End Sub


Private Sub OptCapPar_Click(Index As Integer)
    If Index = 0 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = False
        LblCapPar.Caption = ""
    ElseIf Index = 1 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Linea"
        TxtCapPar.SetFocus
    ElseIf Index = 2 Then
        DTPFecIniPar.Visible = False
        DTPFecFinPar.Visible = False
        LblFecIni.Visible = False
        LblFecFin.Visible = False
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Documento"
        TxtCapPar.SetFocus
    End If

End Sub


Private Sub OptPar_Click(Index As Integer)
        If Index = 0 Then
            LblParos.Caption = "Codigo Paro"
        ElseIf Index = 1 Then
            LblParos.Caption = "Tipo De Paro"
        ElseIf Index = 2 Then
            LblParos.Caption = "Descripcion"
        ElseIf Index = 3 Then
            LblParos.Caption = "Grupo"
        End If
            TxtPar.SetFocus
End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
        
If TabReportes.Tab = 0 Then
        OptPar.Item(0).Value = True
        DBGridConsultas.Visible = True
        DBGridEncabezado.Visible = False
        DBGridDetalle.Visible = False
        FrameTotal.Visible = False
ElseIf TabReportes.Tab = 1 Then
        OptCapPar.Item(0).Value = True
        DTPFecIniPar.Value = Date
        DTPFecFinPar.Value = Date
        DBGridConsultas.Visible = False
        DBGridEncabezado.Visible = True
        DBGridDetalle.Visible = True
        FrameTotal.Visible = True
End If

End Sub




Public Sub Paros()
                       If OptPar.Item(0).Value = True Then
                            DataConsultas.RecordSource = "Select * From Paros Where CodigoParo Like '*" & TxtPar.Text & "*'"
                       ElseIf OptPar.Item(1).Value = True Then
                            DataConsultas.RecordSource = "Select * From Paros Where Tipo = '" & TxtPar.Text & "'"
                       ElseIf OptPar.Item(2).Value = True Then
                            DataConsultas.RecordSource = "Select * From Paros Where DescripcionParo Like '*" & TxtPar.Text & "*'"
                       ElseIf OptPar.Item(3).Value = True Then
                            DataConsultas.RecordSource = "Select * From Paros Where Grupo = '" & TxtPar.Text & "'"
                       End If
                            DataConsultas.Refresh
                            DBGridConsultas.Refresh
                            DBGridConsultas.Columns(1).Width = "4000"

End Sub

Public Sub CapturaParos()
                        If OptCapPar.Item(0).Value = True Then
                                DataEncabezado.RecordSource = "Select * From EncabezadoCapturaParos Where Fecha >= #" & Format(DTPFecIniPar.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFinPar.Value, "mm/dd/yyyy") & "#"
                        ElseIf OptCapPar.Item(1).Value = True Then
                                DataEncabezado.RecordSource = "Select * From EncabezadoCapturaParos Where Fecha >= #" & Format(DTPFecIniPar.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFinPar.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtCapPar.Text & "'"
                        ElseIf OptCapPar.Item(2).Value = True Then
                                DataEncabezado.RecordSource = "Select * From EncabezadoCapturaParos Where Documento = " & TxtCapPar.Text
                        End If
                                DataEncabezado.Refresh
                                DBGridEncabezado.Refresh
                                ColumnasEncabezado
End Sub

Private Sub TxtCapPar_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtCapPar.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
End Sub

Private Sub TxtPar_GotFocus()
        TxtPar.SelStart = 0
        TxtPar.SelLength = Len(TxtPar.Text)
End Sub


Public Sub ColumnasDetalle()
               DBGridDetalle.Columns(0).Width = 1000
               DBGridDetalle.Columns(1).Width = 700
               DBGridDetalle.Columns(2).Width = 700
               DBGridDetalle.Columns(3).Width = 700
               DBGridDetalle.Columns(4).Width = 1000
               DBGridDetalle.Columns(5).Width = 3000
               
            
End Sub

Public Sub ColumnasEncabezado()
               DBGridEncabezado.Columns(0).Width = 1000
               DBGridEncabezado.Columns(1).Width = 1000
               DBGridEncabezado.Columns(2).Width = 700
               DBGridEncabezado.Columns(3).Width = 200
               DBGridEncabezado.Columns(4).Width = 600
               DBGridEncabezado.Columns(5).Width = 600
               DBGridEncabezado.Columns(6).Width = 500
               DBGridEncabezado.Columns(7).Width = 1200
               DBGridEncabezado.Columns(7).NumberFormat = "#,###,##0"
               DBGridEncabezado.Columns(8).Width = 1200
               DBGridEncabezado.Columns(8).NumberFormat = "#,###,##0"
               DBGridEncabezado.Columns(9).Width = 1200
               DBGridEncabezado.Columns(9).NumberFormat = "#,###,##0"
               DBGridEncabezado.Columns(10).Width = 1200
               DBGridEncabezado.Columns(10).NumberFormat = "#,###,##0"
               DBGridEncabezado.Columns(11).Width = 500
               DBGridEncabezado.Columns(12).Width = 500
               DBGridEncabezado.Columns(13).Width = 1000
               
               
End Sub

