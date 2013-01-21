VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Gerencia 
   BackColor       =   &H000000FF&
   Caption         =   "Consulta De Captura De Produccion En Planta"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Gerencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
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
      Height          =   6855
      Left            =   0
      TabIndex        =   25
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
         Picture         =   "Gerencia.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "Gerencia.frx":21BC
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "Gerencia.frx":21D7
         TabIndex        =   33
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin VB.Data DataLineasFichaTecnica 
      Caption         =   "Lineas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data DataGerencia 
      Caption         =   "Gerencia"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
      Caption         =   "Producto No Conforme"
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
      Left            =   2280
      TabIndex        =   21
      Top             =   600
      Width           =   2295
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
      Caption         =   "Producto Conforme"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   360
      Width           =   2055
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
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
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   19
      Top             =   120
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
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
      TabIndex        =   18
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H000000FF&
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
         TabIndex        =   24
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptLinea 
         BackColor       =   &H000000FF&
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
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptGrupo 
         BackColor       =   &H000000FF&
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
         TabIndex        =   22
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
      TabIndex        =   15
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11280
      Picture         =   "Gerencia.frx":2BB1
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSMask.MaskEdBox MskTotalEnvases 
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   8280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTotalTarimas 
      Height          =   255
      Left            =   10080
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin VB.Data DataMes 
      Caption         =   "Mes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data DataDia 
      Caption         =   "Dia"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGridDia 
      Bindings        =   "Gerencia.frx":4C23
      Height          =   3885
      Left            =   120
      OleObjectBlob   =   "Gerencia.frx":4C39
      TabIndex        =   7
      Tag             =   "Produccion Agrupada Por Fecha"
      Top             =   4680
      Width           =   3795
   End
   Begin MSDBGrid.DBGrid DBGridMes 
      Bindings        =   "Gerencia.frx":562A
      Height          =   3210
      Left            =   7800
      OleObjectBlob   =   "Gerencia.frx":5640
      TabIndex        =   6
      Tag             =   "Produccion Agrupada Por Mes"
      Top             =   4680
      Width           =   3975
   End
   Begin MSDBGrid.DBGrid DBGridLineasFichaTecnica 
      Bindings        =   "Gerencia.frx":6025
      Height          =   3735
      Left            =   7800
      OleObjectBlob   =   "Gerencia.frx":604A
      TabIndex        =   5
      Tag             =   "Produccion Agrupada Por Linea"
      Top             =   840
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   61669379
      CurrentDate     =   37248
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
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
      Format          =   61669379
      CurrentDate     =   37248
   End
   Begin MSDBGrid.DBGrid DBGridOrden 
      Bindings        =   "Gerencia.frx":6A42
      Height          =   3855
      Left            =   3960
      OleObjectBlob   =   "Gerencia.frx":6A5A
      TabIndex        =   13
      Tag             =   "Produccion Agrupada Por Orden"
      Top             =   4680
      Width           =   3795
   End
   Begin MSDBGrid.DBGrid DBGridGerencia 
      Bindings        =   "Gerencia.frx":7441
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "Gerencia.frx":745C
      TabIndex        =   0
      Tag             =   "Produccion Agrupada Por Ficha Tecnica"
      Top             =   840
      Width           =   7635
   End
   Begin VB.Label LblLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      TabIndex        =   17
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label LblDescripcion 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   5640
      TabIndex        =   16
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   9120
      TabIndex        =   12
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Tarimas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   9240
      TabIndex        =   11
      Top             =   7920
      Width           =   765
   End
   Begin MSForms.CommandButton CmdGenera 
      Default         =   -1  'True
      Height          =   495
      Left            =   10680
      TabIndex        =   8
      ToolTipText     =   "Generar Datos"
      Top             =   0
      Width           =   495
      BackColor       =   12632256
      PicturePosition =   327683
      Size            =   "873;873"
      Picture         =   "Gerencia.frx":7E52
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "Gerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As Recordset
Dim RBuscaLinea As Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean



Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11

'_______________________________________________________________________________________________________________________
            'GRID DE FICHA TECNICA
            If Opcion.Item(0).Value = True Then
                If OptTodos.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By P.Esp_Tec, F.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Esp_Tec, F.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Esp_Tec, F.Descrip"
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Esp_Tec, F.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Esp_Tec, F.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Esp_Tec, F.Descrip"
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'R') Group By P.Esp_Tec, F.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And (P.Calidad = 'R') Group By P.Esp_Tec, F.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataGerencia.RecordSource = "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas As L, Produccion As P INNER JOIN FichaTecnica As F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And (P.Calidad = 'R') Group By P.Esp_Tec, F.Descrip"
                End If
            End If
            DataGerencia.Refresh
            DBGridGerencia.Refresh
            DBGridGerencia.Columns(3).NumberFormat = "#,###,##0"
            DBGridGerencia.Columns(0).Caption = "Ficha Tecnica"
            DBGridGerencia.Columns(1).Caption = "Descripcion"
            DBGridGerencia.Columns(2).Caption = "Tarimas"
            DBGridGerencia.Columns(3).Caption = "Cantidad"
            DBGridGerencia.Columns(0).Width = "1300"
            DBGridGerencia.Columns(1).Width = "4000"
            DBGridGerencia.Columns(2).Width = "500"
            DBGridGerencia.Columns(3).Width = "900"
            

'_______________________________________________________________________________________________________________________
            'EL GRID DE LINEAS Y FICHA TECNICA
            If Opcion.Item(0).Value = True Then
                If OptTodos.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By P.Linea, L.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Linea, L.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Linea, L.Descrip"
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (Calidad = 'A' OR Calidad = 'I' Or Calidad = 'C') Group By P.Linea, L.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And (Calidad = 'A' OR Calidad = 'I' Or Calidad = 'C') Group By P.Linea, L.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And (Calidad = 'A' OR Calidad = 'I' Or Calidad = 'C') Group By P.Linea, L.Descrip"
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Calidad = 'R' Group By P.Linea, L.Descrip"
                ElseIf OptGrupo.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Linea, L.Descrip"
                ElseIf OptLinea.Value = True Then
                    DataLineasFichaTecnica.RecordSource = "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas as L Inner Join Produccion as P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Linea, L.Descrip"
                End If
            End If
            DataLineasFichaTecnica.Refresh
            DBGridLineasFichaTecnica.Refresh
            
            DBGridLineasFichaTecnica.Columns(3).NumberFormat = "#,###,##0"
            DBGridLineasFichaTecnica.Columns(0).Caption = "Linea"
            DBGridLineasFichaTecnica.Columns(1).Caption = "Descripcion"
            DBGridLineasFichaTecnica.Columns(2).Caption = "Tarimas"
            DBGridLineasFichaTecnica.Columns(3).Caption = "Cantidad"
            DBGridLineasFichaTecnica.Columns(0).Width = "500"
            DBGridLineasFichaTecnica.Columns(1).Width = "1700"
            DBGridLineasFichaTecnica.Columns(2).Width = "500"
            DBGridLineasFichaTecnica.Columns(3).Width = "800"
            
'_______________________________________________________________________________________________________________________
            'EL GRID DE MES
            If Opcion.Item(0).Value = True Then
            'StrSql = "TRANSFORM sum(Produccion.envases) as resultado "
            'StrSql = "SELECT Produccion.Fec_Prd, sum(Produccion.Envases) FROM Produccion "
            'StrSql = StrSql & "GROUP BY Produccion.Fec_Prd PIVOT Month(Produccion.Fec_Prd)"
    
            '    DataMes.RecordSource = StrSql
                If OptTodos.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptGrupo.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas As L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptLinea.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas As L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptGrupo.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas As L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptLinea.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas As L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptGrupo.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                ElseIf OptLinea.Value = True Then
                    DataMes.RecordSource = "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)"
                End If
            End If
            DataMes.Refresh
            DBGridMes.Refresh
            DBGridMes.Columns(3).NumberFormat = "#,###,##0"
            DBGridMes.Columns(0).Caption = "Mes"
            DBGridMes.Columns(1).Caption = "Año"
            DBGridMes.Columns(2).Caption = "Tarimas"
            DBGridMes.Columns(3).Caption = "Cantidad"
            DBGridMes.Columns(0).Width = "500"
            DBGridMes.Columns(1).Width = "500"
            DBGridMes.Columns(2).Width = "900"
            DBGridMes.Columns(3).Width = "1200"
            
            '_______________________________________________________________________________________________________________________
            'PAROS
            
            If OptTodos.Value = True Then
                DataDia.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
            ElseIf OptGrupo.Value = True Then
                DataDia.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
            ElseIf OptLinea.Value = True Then
                DataDia.RecordSource = "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & TxtLinea.Text & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo"
            End If
            DataDia.Refresh
            DBGridDia.Refresh
            
            DBGridDia.Columns(3).NumberFormat = "#,###,##0.00"
            
            DBGridDia.Columns(0).Caption = "Linea"
            DBGridDia.Columns(1).Caption = "Descripcion"
            DBGridDia.Columns(2).Caption = "Tipo"
            DBGridDia.Columns(3).Caption = "Horas"
            
            DBGridDia.Columns(0).Width = "300"
            DBGridDia.Columns(1).Width = "2100"
            DBGridDia.Columns(2).Width = "200"
            DBGridDia.Columns(3).Width = "600"
    
'_______________________________________________________________________________________________________________________
            'EL GRID DE DIA
            If Opcion.Item(0).Value = True Then
                If OptTodos.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By P.Fec_Prd"
                ElseIf OptGrupo.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                ElseIf OptLinea.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Fec_Prd"
                ElseIf OptGrupo.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                ElseIf OptLinea.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' Group By P.Fec_Prd"
                ElseIf OptGrupo.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                ElseIf OptLinea.Value = True Then
                    DataOrden.RecordSource = "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Fec_Prd"
                End If
            End If
            DataOrden.Refresh
            DBGridOrden.Refresh
            DBGridOrden.Columns(2).NumberFormat = "#,###,##0"
            DBGridOrden.Columns(0).NumberFormat = "dd/mm/yyyy"
            DBGridOrden.Columns(0).Caption = "Fecha"
            DBGridOrden.Columns(1).Caption = "Tarimas"
            DBGridOrden.Columns(2).Caption = "Cantidad"
            DBGridOrden.Columns(0).Width = "1000"
            DBGridOrden.Columns(1).Width = "900"
            DBGridOrden.Columns(2).Width = "1200"
'_______________________________________________________________________________________________________________________
            'EL GRID DE ORDEN
            'If Opcion.Item(0).Value = True Then
            '    If OptTodos.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Group By P.Orden"
            '    ElseIf OptGrupo.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Orden"
            '    ElseIf OptLinea.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Orden"
            '    End If
            'ElseIf Opcion.Item(1).Value = True Then
            '    If OptTodos.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Orden"
            '    ElseIf OptGrupo.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Orden"
            '    ElseIf OptLinea.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Orden"
            '    End If
            'ElseIf Opcion.Item(2).Value = True Then
            '    If OptTodos.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' Group By P.Orden"
            '    ElseIf OptGrupo.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By P.Orden"
            '    ElseIf OptLinea.Value = True Then
            '        DataOrden.RecordSource = "Select P.orden, Count(P.Tarima), Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By P.Orden"
            '    End If
            'End If
            'DataOrden.Refresh
            'DBGridOrden.Refresh
            'DBGridOrden.Columns(2).NumberFormat = "#,###,##0"
            'DBGridOrden.Columns(0).Caption = "Orden"
            'DBGridOrden.Columns(1).Caption = "Tarimas"
            'DBGridOrden.Columns(2).Caption = "Cantidad"
            'DBGridOrden.Columns(0).Width = "1000"
            'DBGridOrden.Columns(1).Width = "900"
            'DBGridOrden.Columns(2).Width = "1200"
            
            
'_______________________________________________________________________________________________________________________
            'CUENTA LAS TARIMAS
            If Opcion.Item(0).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C')")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Calidad = 'R'")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Count(P.Tarima) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
            End If
            
            If RTotal.RecordCount > 0 Then
                If Not IsNull(RTotal(0)) Then
                    MskTotalTarimas.Text = RTotal(0)
                Else
                    MskTotalTarimas.Text = "0"
                End If
            Else
                MskTotalTarimas.Text = 0
            End If
'_______________________________________________________________________________________________________________________
            
            'SUMA LOS ENVASES
            If Opcion.Item(0).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C')")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptTodos.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R'")
                ElseIf OptGrupo.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
                ElseIf OptLinea.Value = True Then
                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
                End If
                
            End If
            
            If RTotal.RecordCount > 0 Then
                If Not IsNull(RTotal(0)) Then
                    MskTotalEnvases.Text = RTotal(0)
                Else
                    MskTotalEnvases.Text = "0"
                End If
            Else
                MskTotalEnvases.Text = 0
            End If
            
            
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

Private Sub Form_Load()
            DataGerencia.ConnectionString = GTipoProveedor
            DataLineasFichaTecnica.ConnectionString = GTipoProveedor
            DataMes.ConnectionString = GTipoProveedor
            DataDia.ConnectionString = GTipoProveedor
            DataOrden.ConnectionString = GTipoProveedor
            DataBusqueda.ConnectionString = GTipoProveedor

            DataGerencia.Refresh
            DataLineasFichaTecnica.Refresh
            DataMes.Refresh
            DataDia.Refresh
            DataOrden.Refresh
            DataBusqueda.Refresh
            
            DtpFecIni.Value = Date
            DtpFecFin.Value = Date
            CmdGenera_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
            'DBGridGerencia.Height = Me.ScaleHeight - 5100
            DBGridDia.Height = Me.ScaleHeight - 4700
            DBGridMes.Height = Me.ScaleHeight - 5500
            DBGridOrden.Height = Me.ScaleHeight - 4700
            
            MskTotalTarimas.Move 10000, Me.Height - 1000
            MskTotalEnvases.Move 10000, Me.Height - 700
            Label2.Item(0).Move 8800, Me.Height - 1000
            Label2.Item(1).Move 8800, Me.Height - 700
            If Err <> 0 Then
            End If
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

Private Sub Txtbusqueda_Change()
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Descrip Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Descrip Like '*" & TxtBusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Linea Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas where Linea Like '*" & TxtBusqueda.Text & "*'"
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
                    TxtBusqueda.SetFocus
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
                    TxtBusqueda.SetFocus
        End If
End Sub
