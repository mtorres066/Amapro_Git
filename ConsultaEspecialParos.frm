VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConsultaEspecialParos 
   Caption         =   "Paros De Produccion En Horas"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "ConsultaEspecialParos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Height          =   4815
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   6120
         Picture         =   "ConsultaEspecialParos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "ConsultaEspecialParos.frx":293C
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "ConsultaEspecialParos.frx":2958
         TabIndex        =   25
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.TextBox TxtTotalP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1620
   End
   Begin VB.TextBox TxtTipoNP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1620
   End
   Begin VB.TextBox TxtTipoSP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1620
   End
   Begin VB.TextBox TxtProduccionP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1620
   End
   Begin VB.Data DataConsultas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.OptionButton OptGrupo 
      Caption         =   "Por Grupo"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2500
   End
   Begin VB.TextBox TxtTipoN 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2500
   End
   Begin VB.TextBox TxtTipoS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2500
   End
   Begin VB.TextBox TxtProduccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   6120
      Picture         =   "ConsultaEspecialParos.frx":3347
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdConsultar 
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      Picture         =   "ConsultaEspecialParos.frx":53B9
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Consultar"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Txtlin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton OptLinea 
      Caption         =   "Por Linea"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Todos"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   23789571
      CurrentDate     =   37670
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   23789571
      CurrentDate     =   37670
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   6840
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paros Que No Afectan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paros Que Si Afectan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label LblLinea 
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
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label LblLin 
      Caption         =   "Linea"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "ConsultaEspecialParos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RProduccion As Recordset
Dim RBuscaLinea As Recordset
Dim RTipoS As Recordset
Dim RTipoN As Recordset
Dim RTotal As Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean
Dim VTotal As Single
Dim VProduccion As Single
Dim VTipoS As Single
Dim VTipoN As Single

Dim VTotalP As Single
Dim VProduccionP As Single
Dim VTipoSP As Single
Dim VTipoNP As Single


Private Sub CmdConsultar_Click()
MousePointer = 11
    'TODOS LOS PAROS POR FECHAS
    If OptTodos.Value = True Then
                'SUMA PAROS S
                'DataS.RecordSource = "Select DP.Paro, P.DescripcionParo, sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' Group By DP.Paro, P.DescripcionParo"
                'DataS.Refresh
                'DBGridS.Refresh
                'DBGridS.Columns(1).Width = "4000"
            
                'SUMA PAROS N
                'DataN.RecordSource = "Select DP.Paro, P.DescripcionParo, sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' Group By DP.Paro, P.DescripcionParo"
                'DataN.Refresh
                'DBGridN.Refresh
                'DBGridN.Columns(1).Width = "4000"
                
                'BUSCA LA PRODUCCION
                Set RProduccion = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P'")
                    If RProduccion.RecordCount > 0 Then
                        If IsNull(RProduccion(0)) Then
                            VProduccion = 0
                        Else
                            VProduccion = Format((RProduccion(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VProduccion = 0
                    End If
                'BUSCA LOS TIPO S
                Set RTipoS = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S'")
                    If RTipoS.RecordCount > 0 Then
                        If IsNull(RTipoS(0)) Then
                            VTipoS = 0
                        Else
                            VTipoS = Format((RTipoS(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoS = 0
                    End If
                'BUSCA LOS TIPO N
                Set RTipoN = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N'")
                    If RTipoN.RecordCount > 0 Then
                        If IsNull(RTipoN(0)) Then
                            VTipoN = 0
                        Else
                            VTipoN = Format((RTipoN(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoN = 0
                    End If
                    
                'SUMA TOTALES
                VTotal = VProduccion + VTipoS + VTipoN
                TxtProduccion = VProduccion
                TxtTipoS = VTipoS
                TxtTipoN = VTipoN
                TxtTotal.Text = VTotal
                
                'SACA PORCENTAJEAS
                If VProduccion > 0 Then
                    VProduccionP = Format(VProduccion / VTotal, "#,###,##0.00") * 100
                Else
                    VProduccionP = 0
                End If
                If VTipoS > 0 Then
                    VTipoSP = Format(VTipoS / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoSP = 0
                End If
                If VTipoN > 0 Then
                    VTipoNP = Format(VTipoN / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoNP = 0
                End If
                
                VTotalP = VProduccionP + VTipoSP + VTipoNP
                
                TxtProduccionP = VProduccionP & " %"
                TxtTipoSP = VTipoSP & " %"
                TxtTipoNP = VTipoNP & " %"
                TxtTotalP = VTotalP & " %"
    
    'PAROS POR LINEA
    ElseIf OptLinea.Value = True Then
                
                'BUSCA LA PRODUCCION
                Set RProduccion = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros As P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P'")
                    If RProduccion.RecordCount > 0 Then
                        If IsNull(RProduccion(0)) Then
                            VProduccion = 0
                        Else
                            VProduccion = Format((RProduccion(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VProduccion = 0
                    End If
                
                'BUSCA PAROS TIPO S
                Set RTipoS = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S'")
                    If RTipoS.RecordCount > 0 Then
                        If IsNull(RTipoS(0)) Then
                            VTipoS = 0
                        Else
                            VTipoS = Format((RTipoS(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoS = 0
                    End If
                'BUSCA PAROS TIPO N
                Set RTipoN = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N'")
                    If RTipoN.RecordCount > 0 Then
                        If IsNull(RTipoN(0)) Then
                            VTipoN = 0
                        Else
                            VTipoN = Format((RTipoN(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoN = 0
                    End If
                    
                'SUMA TOTALES
                VTotal = VProduccion + VTipoS + VTipoN
                TxtProduccion = VProduccion
                TxtTipoS = VTipoS
                TxtTipoN = VTipoN
                TxtTotal.Text = VTotal
                
                'SACA PORCENTAJEAS
                If VProduccion > 0 Then
                    VProduccionP = Format(VProduccion / VTotal, "#,###,##0.00") * 100
                Else
                    VProduccionP = 0
                End If
                If VTipoS > 0 Then
                    VTipoSP = Format(VTipoS / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoSP = 0
                End If
                If VTipoN > 0 Then
                    VTipoNP = Format(VTipoN / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoNP = 0
                End If
                
                VTotalP = VProduccionP + VTipoSP + VTipoNP
                
                TxtProduccionP = VProduccionP & " %"
                TxtTipoSP = VTipoSP & " %"
                TxtTipoNP = VTipoNP & " %"
                TxtTotalP = VTotalP & " %"
                
                
    'PAROS POR GRUPO LINEA
    ElseIf OptGrupo.Value = True Then
                
                'BUSCA LA PRODUCCION
                Set RProduccion = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros As P, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P'")
                    If RProduccion.RecordCount > 0 Then
                        If IsNull(RProduccion(0)) Then
                            VProduccion = 0
                        Else
                            VProduccion = Format((RProduccion(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VProduccion = 0
                    End If
                
                'BUSCA PAROS TIPO S
                Set RTipoS = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S'")
                    If RTipoS.RecordCount > 0 Then
                        If IsNull(RTipoS(0)) Then
                            VTipoS = 0
                        Else
                            VTipoS = Format((RTipoS(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoS = 0
                    End If
                'BUSCA PAROS TIPO N
                Set RTipoN = Db.OpenRecordset("Select sum(DP.Minutos) From EncabezadoCapturaParos as EP, DetalleCapturaParos as DP, Paros as P, Lineas as L Where EP.Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N'")
                    If RTipoN.RecordCount > 0 Then
                        If IsNull(RTipoN(0)) Then
                            VTipoN = 0
                        Else
                            VTipoN = Format((RTipoN(0) / 60), "#,###,##0.00")
                        End If
                    Else
                        VTipoN = 0
                    End If
                    
                'SUMA TOTALES
                VTotal = VProduccion + VTipoS + VTipoN
                TxtProduccion = VProduccion
                TxtTipoS = VTipoS
                TxtTipoN = VTipoN
                TxtTotal.Text = VTotal
                
                'SACA PORCENTAJEAS
                If VProduccion > 0 Then
                    VProduccionP = Format(VProduccion / VTotal, "#,###,##0.00") * 100
                Else
                    VProduccionP = 0
                End If
                If VTipoS > 0 Then
                    VTipoSP = Format(VTipoS / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoSP = 0
                End If
                If VTipoN > 0 Then
                    VTipoNP = Format(VTipoN / VTotal, "#,###,##0.00") * 100
                Else
                    VTipoNP = 0
                End If
                
                VTotalP = VProduccionP + VTipoSP + VTipoNP
                
                TxtProduccionP = VProduccionP & " %"
                TxtTipoSP = VTipoSP & " %"
                TxtTipoNP = VTipoNP & " %"
                TxtTotalP = VTotalP & " %"
    End If
    
MousePointer = 0
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
    If BLinea = True Then
        Txtlin.Text = DBGridConsultas.Columns(0)
    ElseIf BGrupo = True Then
        Txtlin.Text = DBGridConsultas.Columns(2)
    End If
        Txtlin.SetFocus
        FrameConsultas.Visible = False
        
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        If BLinea = True Then
            Txtlin.Text = DBGridConsultas.Columns(0)
        ElseIf BGrupo = True Then
            Txtlin.Text = DBGridConsultas.Columns(2)
        End If
            Txtlin.SetFocus
            DBGridConsultas.Visible = False
    End If
End Sub

Private Sub Form_Load()

    DataConsultas.ConnectionString = GTipoProveedor
    DataConsultas.Refresh
    
    'DataN.ConnectionString = GTipoProveedor
    'DataN.Refresh
    
    DTPFecIni.Value = Date
    DTPFecFin.Value = Date
    
    CmdConsultar_Click
End Sub

Private Sub OptGrupo_Click()
    LblLin.Visible = True
    LblLin.Caption = "Grupo"
    LblLinea.Visible = True
    Txtlin.Visible = True
    Txtlin.SetFocus
End Sub

Private Sub OptLinea_Click()
    LblLin.Visible = True
    LblLin.Caption = "Linea"
    LblLinea.Visible = True
    Txtlin.Visible = True
    Txtlin.SetFocus
End Sub

Private Sub OptTodos_Click()
    LblLin.Visible = False
    LblLinea.Visible = False
    Txtlin.Text = ""
    Txtlin.Visible = False
End Sub

Private Sub TxtLin_Change()
    Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
        If RBuscaLinea.RecordCount > 0 Then
            LblLinea.Caption = RBuscaLinea!Descrip
        Else
            LblLinea.Caption = ""
        End If
End Sub

Private Sub Txtlin_DblClick()
    If OptLinea.Value = True Then
        BLinea = True
        BGrupo = False
        DataConsultas.RecordSource = ("Select Linea, Descrip From Lineas")
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        DBGridConsultas.Columns(0).Width = "500"
        DBGridConsultas.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
    ElseIf OptGrupo.Value = True Then
        BLinea = False
        BGrupo = True
        DataConsultas.RecordSource = ("Select Linea, Descrip, Grupo From Lineas")
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        DBGridConsultas.Columns(0).Width = "500"
        DBGridConsultas.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
    End If
End Sub

Private Sub TxtLin_GotFocus()
    Txtlin.SelStart = 0
    Txtlin.SelLength = Len(Txtlin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
            If OptLinea.Value = True Then
                BLinea = True
                BGrupo = False
                DataConsultas.RecordSource = ("Select Linea, Descrip From Lineas")
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                DBGridConsultas.Columns(0).Width = "500"
                DBGridConsultas.Columns(1).Width = "4000"
                FrameConsultas.Visible = True
                DBGridConsultas.SetFocus
            ElseIf OptGrupo.Value = True Then
                BLinea = False
                BGrupo = True
                DataConsultas.RecordSource = ("Select Linea, Descrip, Grupo From Lineas")
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                DBGridConsultas.Columns(0).Width = "500"
                DBGridConsultas.Columns(1).Width = "4000"
                FrameConsultas.Visible = True
                DBGridConsultas.SetFocus
            End If
    End If
End Sub

