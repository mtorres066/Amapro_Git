VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConsultaProgramacionProduccion 
   Caption         =   "Consulta Programacion De La Produccion"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrReportes 
      Left            =   5160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton CmdImprimir 
      Height          =   495
      Left            =   9600
      Picture         =   "ConsultaProgramacionProduccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Default         =   -1  'True
      Height          =   495
      Left            =   11160
      Picture         =   "ConsultaProgramacionProduccion.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdAceptar 
      Height          =   495
      Left            =   10440
      Picture         =   "ConsultaProgramacionProduccion.frx":21BC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Consultar"
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64880641
      CurrentDate     =   37686
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64880641
      CurrentDate     =   37686
   End
   Begin VB.Data DataProgramacion 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGridProgramacion 
      Bindings        =   "ConsultaProgramacionProduccion.frx":422E
      Height          =   7815
      Left            =   120
      OleObjectBlob   =   "ConsultaProgramacionProduccion.frx":424D
      TabIndex        =   0
      Top             =   720
      Width           =   11655
   End
   Begin VB.Label Label1 
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
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "ConsultaProgramacionProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VDia As String
Dim VMes As String
Dim VAño As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String


Private Sub CmdAceptar_Click()
    DataProgramacion.RecordSource = "Select P.Fecha, P.Linea, L.Descrip, P.Turno, P.FichaTecnica, F.Descrip, P.Cantidad From ProgramacionProduccion as P, FichaTecnica as F, Lineas as L Where P.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And P.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And P.FichaTecnica = F.Esp_Tec Order By P.Fecha, P.Linea, P.Turno, P.FichaTecnica"
    DataProgramacion.Refresh
    DBGridProgramacion.Refresh
    
    DBGridProgramacion.Columns(0).Width = "1000"
    DBGridProgramacion.Columns(1).Width = "500"
    DBGridProgramacion.Columns(2).Width = "3000"
    DBGridProgramacion.Columns(3).Width = "300"
    DBGridProgramacion.Columns(4).Width = "1500"
    DBGridProgramacion.Columns(5).Width = "3700"
    DBGridProgramacion.Columns(6).Width = "1300"
    DBGridProgramacion.Columns(6).NumberFormat = "#,###,##0.00"
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
            VDia = Day(DtpFecIni.Value)
            VMes = Month(DtpFecIni.Value)
            VAño = Year(DtpFecIni.Value)
            VDia2 = Day(DtpFecFin.Value)
            VMes2 = Month(DtpFecFin.Value)
            VAño2 = Year(DtpFecFin.Value)
                 
            CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpFecFin.Value, "dd/mm/yyyy") & "'"
            CrReportes.SelectionFormula = "{ProgramacionProduccion.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                    
MousePointer = 0
            CrReportes.ReportFileName = App.Path & "\ProgramacionProduccion.rpt"
            CrReportes.Action = 0
            
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DataProgramacion.Connect = GConnect
    DataProgramacion.DatabaseName = BasedeDatos
    DtpFecIni.Value = Date
    DtpFecFin.Value = Date
    CmdAceptar_Click
End Sub
