VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesDeOrdenes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes De Ordenes"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "ReportesDeOrdenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   9615
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5655
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9975
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
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   2292
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   615
         Left            =   8760
         Picture         =   "ReportesDeOrdenes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblbusqueda 
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
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   516
      End
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
      Height          =   1935
      Left            =   8400
      Picture         =   "ReportesDeOrdenes.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
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
      Height          =   1935
      Left            =   8400
      Picture         =   "ReportesDeOrdenes.frx":2DE6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   1058
      BackColor       =   12632256
      TabCaption(0)   =   "Ordenes De Produccion"
      TabPicture(0)   =   "ReportesDeOrdenes.frx":3530
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblOrdFec(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblOrdFec(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblOrden2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblOrden"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblBodega"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblBodega2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTPOrdFecIni"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPOrdFecFin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameOrden"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameOrden3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtOrden"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameBodegas"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtBodega"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox TxtBodega 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   5520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Tipo De Bodega"
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
         Height          =   1092
         Left            =   6240
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1692
         Begin VB.OptionButton OptBod 
            Caption         =   "Todas"
            Height          =   192
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   852
         End
         Begin VB.OptionButton OptBod 
            Caption         =   "No Conforme"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1332
         End
         Begin VB.OptionButton OptBod 
            Caption         =   "De Proceso"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1332
         End
      End
      Begin VB.TextBox TxtOrden 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   2400
         TabIndex        =   20
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Frame FrameOrden3 
         Caption         =   "Tipo De Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   972
         Left            =   6240
         TabIndex        =   12
         Top             =   600
         Width           =   1692
         Begin VB.OptionButton OptOrdAbi 
            Caption         =   "Abiertas"
            Height          =   192
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   972
         End
         Begin VB.OptionButton OptOrdCer 
            Caption         =   "Cerradas"
            Height          =   192
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   972
         End
         Begin VB.OptionButton OptOrdTod 
            Caption         =   "Todas"
            Height          =   192
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   972
         End
      End
      Begin VB.Frame FrameOrden 
         Caption         =   "Tipo De Reporte"
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
         Height          =   4095
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   5895
         Begin VB.OptionButton OptOrd 
            Caption         =   "Detalle De Consumos Por Fechas y Orden"
            Height          =   192
            Index           =   11
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Ventas Por Fechas Formato Resumen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   192
            Index           =   10
            Left            =   120
            TabIndex        =   37
            Top             =   3840
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Ventas Por Fechas Formato Detalle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   192
            Index           =   9
            Left            =   120
            TabIndex        =   36
            Top             =   3600
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Inventario Por Fechas"
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
            Height          =   192
            Index           =   8
            Left            =   120
            TabIndex        =   35
            Top             =   3240
            Width           =   2415
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Resumen De Traslados Por Fechas y Bodega Entrada y Orden"
            Height          =   192
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   4935
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Resumen De Traslados Por Fechas y Bodega Salida y Orden"
            Height          =   192
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   2040
            Width           =   4695
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Resumen De Traslados Por Fechas y Orden"
            Height          =   192
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Listado Detalle De Ordenes"
            Height          =   192
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   2880
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Listado Resumen De Ordenes"
            Height          =   192
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   2640
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Resumen De Consumos Por Fechas y Orden"
            Height          =   192
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Resumen De Produccion Por Fechas y Orden"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3612
         End
         Begin VB.OptionButton OptOrd 
            Caption         =   "Cedula De La Orden"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1812
         End
      End
      Begin MSComCtl2.DTPicker DTPOrdFecFin 
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51249153
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker DTPOrdFecIni 
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51249153
         CurrentDate     =   37722
      End
      Begin VB.Label LblBodega2 
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
         Left            =   3720
         TabIndex        =   34
         Top             =   5520
         Width           =   4215
      End
      Begin VB.Label LblBodega 
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
         Left            =   1560
         TabIndex        =   33
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LblOrden 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label LblOrden2 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         TabIndex        =   22
         Top             =   5880
         Width           =   4215
      End
      Begin VB.Label LblOrdFec 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   1080
         TabIndex        =   21
         Top             =   4800
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label LblOrdFec 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   5160
         Visible         =   0   'False
         Width           =   1005
      End
   End
End
Attribute VB_Name = "ReportesDeOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Dim myReport As CRAXDRT.Report

'Dim mySubReport As SubreportObject
            
Dim VDia As String
Dim VDia2 As String
Dim VMes As String
Dim VMes2 As String
Dim VAño As String
Dim VAño2 As String

Dim RBusqueda As New ADODB.Recordset

'EFICIENCIA
Dim RTiempoProgramadoD As New ADODB.Recordset
Dim VTiempoProgramadoD As Single

Dim RTiempoProgramadoN As New ADODB.Recordset
Dim VTiempoProgramadoN As Single

'PAROS QUE NO AFECTAN LA PRODUCCION
Dim RBuscaParosNoAfectanD As New ADODB.Recordset
Dim VParosND As Single

Dim RBuscaParosNoAfectanN As New ADODB.Recordset
Dim VParosNN As Single

'PARO QUE SI AFECTAN LA PRODUCCION
Dim RBuscaParosSiAfectanD As New ADODB.Recordset
Dim VParosSD As Single

Dim RBuscaParosSiAfectanN As New ADODB.Recordset
Dim VParosSN As Single

'PRODUCCION
Dim RBuscaProduccionD As New ADODB.Recordset
Dim VProduccionD As Single

Dim RBuscaProduccionN As New ADODB.Recordset
Dim VProduccionN As Single


'TIEMPO REAL DE PRODUCCION
Dim VTiempoRealProducidoD As Single
Dim VTiempoRealProducidoN As Single

Dim VHorasProducidasDN As Single
Dim VParosDN As Single

Dim VPCD As Single
Dim VPNCD As Single
Dim VPDD As Single

Dim VPCN As Single
Dim VPNCN As Single
Dim VPDN As Single

Dim RProduccion As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset

Dim VTotalProduccion As Single
Dim VTotalProduccionD As Single
Dim VTotalProduccionN As Single

Dim VVelocidadPromedio As Integer
Dim VVelocidadPromedioD As Integer
Dim VVelocidadPromedioN As Integer

Dim VVelocidadTeoricaLinea As Integer
Dim VVelocidadRealLinea As Integer

'CALCULOS DE EFICIENCIA
Dim VVelocidadTeoricaDia As Integer
Dim VVelocidadTeoricaNoche As Integer
Dim VVelocidadRealDia As Integer
Dim VVelocidadRealNoche As Integer

Dim VFactor1D As Single
Dim VFactor2D As Single
Dim VFactor3D As Single
Dim VFactor4D As Single
Dim VFactor5D As Single

Dim VFactor1N As Single
Dim VFactor2N As Single
Dim VFactor3N As Single
Dim VFactor4N As Single
Dim VFactor5N As Single

Dim VFactor1DN As Single
Dim VFactor2DN As Single
Dim VFactor3DN As Single
Dim VFactor4DN As Single
Dim VFactor5DN As Single

Dim VFactor1TDN As Single
Dim VFactor2TDN As Single
Dim VFactor3TDN As Single
Dim VFactor4TDN As Single
Dim VFactor5TDN As Single

Dim RLineas As New ADODB.Recordset

Dim VEficienciaRealD As Single
Dim VEficienciaRealN As Single

Dim RReporteEficiencia As New ADODB.Recordset

Dim VLinea As String
Dim VFechaInicial As Date
Dim VFechaFinal As Date

Dim RSeleccionaLineas As New ADODB.Recordset

Dim VPorcentajeLinea As Single
Dim VPorcentajeRechazo As Single
Dim VPorcentajeDesperdicio As Single
Dim VPorcentajeRechazoD As Single
Dim VPorcentajeDesperdicioD As Single
Dim VPorcentajeRechazoN As Single
Dim VPorcentajeDesperdicioN As Single



Dim Cont As Integer
'Dim VFactorUno As Double
Dim RBuscaDescripcionLinea As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset
Dim RBuscaEmpleado As New ADODB.Recordset

'VARIABLES PARA BUSQUEDA DE DATOS
Dim BGrupos As Boolean
Dim BParos As Boolean
Dim BGrupos2 As Boolean
Dim BEficiencia As Boolean
Dim BEficiencia2 As Boolean
Dim BGrupoParo As Boolean
Dim BCliente As Boolean

Dim RBuscaOrden As New ADODB.Recordset
Dim RCapturaParos As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset

Dim BMenorHorasD As Boolean
Dim BMenorHorasN As Boolean

Dim VGrupoDia As String
Dim VGrupoNoche As String

Dim RBuscaBodega As New ADODB.Recordset
Dim BBodega As Boolean
Dim BOrden As Boolean


Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11


  
  If TabReportes.Tab = 0 Then
                                Ordenes
                                
  End If
  
'********************************************************************************************************************************************************************************************
                MousePointer = 0
                FrmReporte.Show
                
                
                If Err > 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                End If
  
End Sub

Private Sub CmdSalida_Click()
    Unload Me
    
End Sub


Private Sub CmdSalir_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DBGridBusqueda_DblClick()
        If BOrden = True Then
            TxtOrden.Text = DBGridBusqueda.Columns(0).Text
        ElseIf BBodega = True Then
            TxtBodega.Text = DBGridBusqueda.Columns(0).Text
        End If
            FrameBusqueda.Visible = False
            
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If BOrden = True Then
            TxtOrden.Text = DBGridBusqueda.Columns(0).Text
        ElseIf BBodega = True Then
            TxtBodega.Text = DBGridBusqueda.Columns(0).Text
        End If
            FrameBusqueda.Visible = False
            
End Sub

Private Sub Form_Load()
            
            DTPOrdFecIni.Value = Date
            DTPOrdFecFin.Value = Date
            
            'Set myReport = New CrystalReport1
            'Set mySubReport = myReport.Orders
            
            If GInvVenRepEje = True Then
                OptOrd.Item(9).Visible = True
                OptOrd.Item(10).Visible = True
            Else
                OptOrd.Item(9).Visible = False
                OptOrd.Item(10).Visible = False
            End If

            
End Sub

Private Sub OptOrd_Click(Index As Integer)
        If Index = 0 Or Index = 4 Or Index = 5 Then
                DTPOrdFecIni.Visible = False
                DTPOrdFecFin.Visible = False
                LblOrdFec.Item(0).Visible = False
                LblOrdFec.Item(1).Visible = False
        Else
                DTPOrdFecIni.Visible = True
                DTPOrdFecFin.Visible = True
                LblOrdFec.Item(0).Visible = True
                LblOrdFec.Item(1).Visible = True
                LblOrdFec.Item(0).Caption = "Fecha Inicial"
                LblOrdFec.Item(1).Caption = "Fecha Final"
        End If
                TxtOrden.SetFocus
           
        'TIPO DE BODEGA
        If Index = 3 Or Index = 6 Or Index = 7 Then
                FrameBodegas.Visible = True
        Else
                FrameBodegas.Visible = False
        End If
        
        'TRASLADOS
        If Index = 6 Or Index = 7 Then
                LblBodega.Visible = True
                TxtBodega.Visible = True
                TxtBodega.SetFocus
        Else
                LblBodega.Visible = False
                TxtBodega.Visible = False
        End If
        
        'INVENTARIO
        If Index = 8 Then
                DTPOrdFecIni.Visible = True
                DTPOrdFecFin.Visible = True
                LblOrdFec.Item(0).Visible = True
                LblOrdFec.Item(1).Visible = True
                LblOrdFec.Item(0).Caption = "Fecha Inicial"
                LblOrdFec.Item(1).Caption = "Fecha Final"
        End If
        
        'VENTAS
        If Index = 9 Or Index = 10 Then
                DTPOrdFecIni.Visible = True
                DTPOrdFecFin.Visible = True
                LblOrdFec.Item(0).Visible = True
                LblOrdFec.Item(1).Visible = True
                LblOrdFec.Item(0).Caption = "Fecha Operacion Inicial"
                LblOrdFec.Item(1).Caption = "Fecha Operacion Final"
        End If
End Sub

Private Sub TxtBodega_Change()
        Set RBuscaBodega = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodega.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasMateriaPrima Where UPPER(CodigoBodega) = '" & UCase(TxtBodega.Text) & "'")
            End If
            If RBuscaBodega.RecordCount > 0 Then
                    LblBodega2.Caption = RBuscaBodega!Descripcion
            Else
                    LblBodega2.Caption = ""
            End If
            
End Sub

Private Sub TxtBodega_DblClick()
                    BBodega = True
                    BOrden = False
                    FrameBusqueda.Visible = True
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasMateriaPrima")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    TxtBusqueda.SetFocus

End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
                If KeyAscii = 43 Then
                    BBodega = True
                    BOrden = False
                    FrameBusqueda.Visible = True
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasMateriaPrima")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    TxtBusqueda.SetFocus
                End If
End Sub

Private Sub TxtBusqueda_Change()
    If BOrden = True Then
            FrameBusqueda.Visible = True
            Set RBusqueda = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where EO.FichaTecnica = F.Esp_Tec And EO.Documento Like '%" & TxtBusqueda.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where UPPER(EO.FichaTecnica) = UPPER(F.Esp_Tec) And UPPER(EO.Documento) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(2).Width = "4000"
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

Private Sub TxtOrden_Change()
    'ORDEN
    If OptOrd.Item(0).Value = True Then
        Set RBuscaOrden = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaOrden, "Select EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where EO.Documento = '" & TxtOrden.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaOrden, "Select EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where UPPER(EO.Documento) = '" & UCase(TxtOrden.Text) & "'")
            End If
            If RBuscaOrden.RecordCount > 0 Then
                LblOrden2.Caption = RBuscaOrden(0) & Space(5) & RBuscaOrden(1)
            Else
                LblOrden2.Caption = ""
            End If
    End If
End Sub

Private Sub TxtOrden_DblClick()
                        BOrden = True
                        BBodega = False
                        FrameBusqueda.Visible = True
                        Set RBusqueda = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where EO.FichaTecnica = F.Esp_Tec")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where UPPER(EO.FichaTecnica) = UPPER(F.Esp_Tec)")
                            End If
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(2).Width = "4000"
                        TxtBusqueda.SetFocus
End Sub

Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                        SendKeys "{tab}"
                End If
                
                
                If KeyAscii = 43 Then
                        BOrden = True
                        BBodega = False
                        FrameBusqueda.Visible = True
                        Set RBusqueda = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where EO.FichaTecnica = F.Esp_Tec")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select EO.Documento, EO.FichaTecnica, F.Descrip From EncabezadoOrdenProduccion EO, FichaTecnica F Where UPPER(EO.FichaTecnica) = UPPER(F.Esp_Tec)")
                            End If
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(2).Width = "4000"
                        TxtBusqueda.SetFocus
               End If
End Sub


Public Sub Ordenes()
                        
                        VDia = Day(DTPOrdFecIni.Value)
                        VMes = Month(DTPOrdFecIni.Value)
                        VAño = Year(DTPOrdFecIni.Value)
                        VDia2 = Day(DTPOrdFecFin.Value)
                        VMes2 = Month(DTPOrdFecFin.Value)
                        VAño2 = Year(DTPOrdFecFin.Value)
                                    
                                    'TITULO
                                    If OptOrd.Item(0).Value = True Then
                                    Else
                                                       GTituloReporte = "Desde " & DTPOrdFecIni.Value & " Hasta " & DTPOrdFecFin.Value
                                    End If
                                    'CEDULA
                                    If OptOrd.Item(0).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoOrdenProduccion.Documento} = '" & TxtOrden.Text & "'"
                                    'PRODUCCION
                                    ElseIf OptOrd.Item(1).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Documento} = {DetalleProduccionPorOrden.Documento} And {DetalleProduccionPorOrden.Orden} Like '" & TxtOrden.Text & "*'"
                                    'CONSUMOS
                                    ElseIf OptOrd.Item(2).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleConsumoMateriaPrima.Orden} Like '" & TxtOrden.Text & "*'"
                                    'CONSUMOS DETALLE
                                    ElseIf OptOrd.Item(11).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleConsumoMateriaPrima.Orden} Like '" & TxtOrden.Text & "*'"
                                    'TRASLADOS
                                    ElseIf OptOrd.Item(3).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.Orden} Like '" & TxtOrden.Text & "*'"
                                    'LISTADO RESUMEN
                                    ElseIf OptOrd.Item(4).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoOrdenProduccion.Documento} Like '" & TxtOrden.Text & "*'"
                                    'LISTADO DETALLE
                                    ElseIf OptOrd.Item(5).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoOrdenProduccion.Documento} Like '" & TxtOrden.Text & "*'"
                                                       
                                    ElseIf OptOrd.Item(6).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoTrasladosMateriaPrim.BodegaSalida} = '" & TxtBodega.Text & "' And {DetalleTrasladosMateriaPrimaP.Orden} Like '" & TxtOrden.Text & "*'"
                                    ElseIf OptOrd.Item(7).Value = True Then
                                                       GCriteriaReporte = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.BodegaEntrada} = '" & TxtBodega.Text & "' And {DetalleTrasladosMateriaPrimaP.Orden} Like '" & TxtOrden.Text & "*'"
                                    'INVENTARIO
                                    ElseIf OptOrd.Item(8).Value = True Then
                                                       GCriteriaReporte = "{Inventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                    'VENTAS DETALLE
                                    ElseIf OptOrd.Item(9).Value = True Then
                                                       GCriteriaReporte = "{VentasDetalle.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                    'VENTAS RESUMEN
                                    ElseIf OptOrd.Item(10).Value = True Then
                                                       GCriteriaReporte = "{VentasDetalle.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                    
                                    End If
                                    
                                    'TIPO DE BODEGA
                                    If OptOrd.Item(3).Value = True Or OptOrd.Item(6).Value = True Or OptOrd.Item(7).Value = True Then
                                            If OptBod.Item(0).Value = True Then
                                                GCriteriaReporte = GCriteriaReporte & " And {EncabezadoTrasladosMateriaPrim.BodegaSalida} = {BodegasMateriaPrima2.CodigoBodega} And {BodegasMateriaPrima2.EsBodegadeProceso} = true"
                                            ElseIf OptBod.Item(1).Value = True Then
                                                GCriteriaReporte = GCriteriaReporte & " And {EncabezadoTrasladosMateriaPrim.BodegaSalida} = {BodegasMateriaPrima2.CodigoBodega} And {BodegasMateriaPrima2.EsBodegadeNoConforme} = true"
                                            Else
                                            End If
                                    
                                    End If
                                                
                                    'ABIERTAS
                                    If OptOrdAbi.Value = True Then
                                        GCriteriaReporte = GCriteriaReporte & " And {EncabezadoOrdenProduccion.Estado} = 'ABIERTA'"
                                    'CERRADAS
                                    ElseIf OptOrdCer.Value = True Then
                                        GCriteriaReporte = GCriteriaReporte & " And {EncabezadoOrdenProduccion.Estado} = 'CERRADA'"
                                    'TODOS
                                    ElseIf OptOrdTod.Value = True Then
                                        'NO HACE NADA
                                    End If
                                     
                                    'DETALLE
                                    If OptOrd.Item(0).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionDetalle.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionDetalleO.rpt"
                                        End If
                                    'PRODUCCION
                                    ElseIf OptOrd.Item(1).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionResumenConProduccion.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionResumenConProduccionO.rpt"
                                        End If
                                    'CONSUMOS
                                    ElseIf OptOrd.Item(2).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionResumenConConsumos.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionResumenConConsumosO.rpt"
                                        End If
                                    'TRASLADOS
                                    ElseIf OptOrd.Item(3).Value = True Or OptOrd.Item(6).Value = True Or OptOrd.Item(7).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionResumenConTraslados.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionResumenConTrasladosO.rpt"
                                        End If
                                    'LISTADO RESUMEN
                                    ElseIf OptOrd.Item(4).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionListadoResumen.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionListadoResumenO.rpt"
                                        End If
                                    'LISTADO DETALLE
                                    ElseIf OptOrd.Item(5).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionListadoDetalle.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionListadoDetalleO.rpt"
                                        End If
                                    'INVENTARIO
                                    ElseIf OptOrd.Item(8).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "Inventario.rpt"
                                        Else
                                            GNombreReporte = "InventarioO.rpt"
                                        End If
                                    'VENTAS DETALLE
                                    ElseIf OptOrd.Item(9).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "Ventas.rpt"
                                        Else
                                            GNombreReporte = "VentasO.rpt"
                                        End If
                                    'VENTAS RESUMEN
                                    ElseIf OptOrd.Item(10).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "VentasResumen.rpt"
                                        Else
                                            GNombreReporte = "VentasResumenO.rpt"
                                        End If
                                    'CONSUMOS DETALLE
                                    ElseIf OptOrd.Item(11).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "OrdenProduccionDetalleConConsumos.rpt"
                                        Else
                                            GNombreReporte = "OrdenProduccionDetalleConConsumosO.rpt"
                                        End If
                                    End If
          
End Sub
