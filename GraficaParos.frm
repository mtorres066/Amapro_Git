VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GraficaParos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Grafica De 10 Paros Mas Altos"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   Icon            =   "GraficaParos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
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
      Height          =   6855
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "GraficaParos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
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
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "GraficaParos.frx":237C
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "GraficaParos.frx":2397
         TabIndex        =   26
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin VB.OptionButton OptHoras 
      Caption         =   "Horas"
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
      Height          =   195
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton OptMinutos 
      Caption         =   "Minutos"
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
      Height          =   195
      Left            =   6840
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CDDialogo 
      Left            =   10680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp;JPEG"
      DialogTitle     =   "Grabar Grafica"
      Filter          =   "Pictures (*.bmp)|*.bmp"
      FilterIndex     =   3
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      MaxLength       =   2
      TabIndex        =   3
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Width           =   495
   End
   Begin TabDlg.SSTab tabGrafica 
      Height          =   7695
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   13568
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Todos Los Paros"
      TabPicture(0)   =   "GraficaParos.frx":2D71
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grafica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paros Que Si Afectan 'S'"
      TabPicture(1)   =   "GraficaParos.frx":308B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GraficaS"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Paros Que No Afectan 'N'"
      TabPicture(2)   =   "GraficaParos.frx":33A5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GraficaN"
      Tab(2).ControlCount=   1
      Begin MSChart20Lib.MSChart GraficaS 
         Height          =   6855
         Left            =   -74880
         OleObjectBlob   =   "GraficaParos.frx":36BF
         TabIndex        =   16
         Top             =   720
         Width           =   11535
      End
      Begin MSChart20Lib.MSChart GraficaN 
         Height          =   6855
         Left            =   -74880
         OleObjectBlob   =   "GraficaParos.frx":4F0F
         TabIndex        =   17
         Top             =   720
         Width           =   11535
      End
      Begin MSChart20Lib.MSChart Grafica 
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "GraficaParos.frx":675F
         TabIndex        =   15
         Top             =   840
         Width           =   11535
      End
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   11280
      Picture         =   "GraficaParos.frx":810E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdGrabar 
      Height          =   495
      Left            =   8880
      Picture         =   "GraficaParos.frx":A180
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar Grafica"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdCopiar 
      Height          =   495
      Left            =   10080
      Picture         =   "GraficaParos.frx":A6B2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Copiar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdImprimirGrafica 
      Height          =   495
      Left            =   9480
      Picture         =   "GraficaParos.frx":ABE4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir Grafica"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox CboVerGra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "GraficaParos.frx":B116
      Left            =   1200
      List            =   "GraficaParos.frx":B141
      TabIndex        =   0
      Text            =   "2dBar"
      Top             =   0
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   445
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   62849027
      CurrentDate     =   37153
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   445
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   62849027
      CurrentDate     =   37153
   End
   Begin VB.CommandButton CmdGenerar 
      Default         =   -1  'True
      Height          =   495
      Left            =   10680
      Picture         =   "GraficaParos.frx":B1B8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Generar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LblLin 
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
      Left            =   1200
      TabIndex        =   19
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Linea"
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
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   6480
      TabIndex        =   13
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Vistas De Grafica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "GraficaParos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RCuentaParos As Recordset
Dim RParos As Recordset
Dim RParosS As Recordset
Dim RParosN As Recordset
Dim RBuscaDescripcionParo As Recordset

Dim CantidadDeColumnas As Integer
Dim cont As Integer

Dim RLineas As Recordset
Dim VEtiqueta As String

Dim TotalT As Double
Dim TotalN As Double
Dim TotalS As Double
Dim RGraficaDeParos As Recordset

'Dim cnn As New ADODB.Connection
'Dim rst As New ADODB.Recordset
'Dim strProveedor As String
'Dim strOrigendatos As String
'Dim strSQL As String


Private Sub CboVerGra_Click()
If CboVerGra.ListIndex = 0 Then
            Grafica.chartType = VtChChartType2dArea
            GraficaS.chartType = VtChChartType2dArea
            GraficaN.chartType = VtChChartType2dArea
ElseIf CboVerGra.ListIndex = 1 Then
            Grafica.chartType = VtChChartType2dBar
            GraficaS.chartType = VtChChartType2dBar
            GraficaN.chartType = VtChChartType2dBar
ElseIf CboVerGra.ListIndex = 2 Then
            Grafica.chartType = VtChChartType2dCombination
            GraficaS.chartType = VtChChartType2dCombination
            GraficaN.chartType = VtChChartType2dCombination
ElseIf CboVerGra.ListIndex = 3 Then
            Grafica.chartType = VtChChartType2dLine
            GraficaS.chartType = VtChChartType2dLine
            GraficaN.chartType = VtChChartType2dLine
ElseIf CboVerGra.ListIndex = 4 Then
            Grafica.chartType = VtChChartType2dPie
            GraficaS.chartType = VtChChartType2dPie
            GraficaN.chartType = VtChChartType2dPie
ElseIf CboVerGra.ListIndex = 5 Then
            Grafica.chartType = VtChChartType2dStep
            GraficaS.chartType = VtChChartType2dStep
            GraficaN.chartType = VtChChartType2dStep
ElseIf CboVerGra.ListIndex = 6 Then
            Grafica.chartType = VtChChartType2dXY
            GraficaS.chartType = VtChChartType2dXY
            GraficaN.chartType = VtChChartType2dXY
ElseIf CboVerGra.ListIndex = 7 Then
            Grafica.chartType = VtChChartType3dArea
            GraficaS.chartType = VtChChartType3dArea
            GraficaN.chartType = VtChChartType3dArea
ElseIf CboVerGra.ListIndex = 8 Then
            Grafica.chartType = VtChChartType3dBar
            GraficaS.chartType = VtChChartType3dBar
            GraficaN.chartType = VtChChartType3dBar
ElseIf CboVerGra.ListIndex = 9 Then
            Grafica.chartType = VtChChartType3dCombination
            GraficaS.chartType = VtChChartType3dCombination
            GraficaN.chartType = VtChChartType3dCombination
ElseIf CboVerGra.ListIndex = 10 Then
            Grafica.chartType = VtChChartType3dLine
            GraficaS.chartType = VtChChartType3dLine
            GraficaN.chartType = VtChChartType3dLine
ElseIf CboVerGra.ListIndex = 11 Then
            Grafica.chartType = VtChChartType3dStep
            GraficaS.chartType = VtChChartType3dStep
            GraficaN.chartType = VtChChartType3dStep
End If

End Sub

Private Sub CmdCopiar_Click()
    If tabGrafica.Tab = 0 Then
            Grafica.EditCopy
    ElseIf tabGrafica.Tab = 1 Then
            GraficaS.EditCopy
    ElseIf tabGrafica.Tab = 2 Then
            GraficaN.EditCopy
    End If
        

End Sub

Private Sub CmdGenerar_Click()
MousePointer = 11
On Error Resume Next
            
            'BUSCA SI EXISTE LA LINEA
            Set RLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RLineas.RecordCount > 0 Then
            Else
                MsgBox "Linea No Existe", vbOKOnly + vbInformation, "Informacion"
                Txtlin.SetFocus
                Exit Sub
                
            End If

            
            'strSQL = "Select paro, sum(minutos) From DetalleCapturaParos Group By Paro"
            'strProveedor = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & BasedeDatos & ";Mode=ReadWrite"
            'cnn.Open strProveedor
            'rst.Open strSQL, cnn, adOpenStatic
            'Set grafica.DataSource = rst
            'Set DataGrid1.DataSource = rst
            'If Err <> 0 Then
            '    MsgBox Err.Number & Err.Description
            'End If
            'rst.Close
            'cnn.Close
            
            Db.Execute ("Delete * From Graficadeparos")
            
            Set RGraficaDeParos = Db.OpenRecordset("Select * From GraficaDeParos")
            
'TODOS LOS PAROS______________________________________________________________________________________________________
            
            CantidadDeColumnas = 0
            cont = 1
            
                        'CUENTA CUANDOS GRUPOS SE CREAN PARA CALCULAR LAS COLUMNAS DEL GRAFICA
                        Set RCuentaParos = Db.OpenRecordset("Select top 10 Count(*) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento Group By DP.Paro")
                            If RCuentaParos.RecordCount > 0 Then
                                Do Until RCuentaParos.EOF
                                            CantidadDeColumnas = CantidadDeColumnas + 1
                                        RCuentaParos.MoveNext
                                Loop
                            Else
                                CantidadDeColumnas = 0
                            End If
            
            
                            'AGRUPA LOS PAROS POR CODIGO DE PARO
                            Set RParos = Db.OpenRecordset("Select Top 10 Paro, Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento Group By DP.Paro Order By DP.Paro")
                            
                            Do Until RParos.EOF
                                    'AGREGA DATOS A LA BASE DE DATOS
                                    RGraficaDeParos.AddNew
                                            RGraficaDeParos!CodigoParo = RParos(0)
                                            'SI SELECCIONA LA OPCION EN MINUTOS
                                            If OptMinutos.Value = True Then
                                                RGraficaDeParos!Minutos = RParos(1)
                                            Else
                                            'SI SELECCIONA LA OPCION EN PULGADAS
                                                RGraficaDeParos!Minutos = Format((RParos(1) / 60), "#,###,##0.00")
                                            End If
                                            RGraficaDeParos!Tipo = "T"
                                    RGraficaDeParos.Update
                                RParos.MoveNext
                            Loop
                                    
                            'SACA TODOS LOS PAROS DE LA BASE DE DATOS ORDENADOS POR MINUTOS
                            Set RParos = Db.OpenRecordset("Select * from GraficaDeParos Where Tipo = 'T' Order by Minutos")
                            
                            
                            'CREA LA CANTIDAD DE COLUMNAS QUE VA A TENER LA GRAFICA
                            Grafica.ColumnCount = CantidadDeColumnas
                                        
                            Do Until RParos.EOF
                                        Grafica.Column = cont
                                        Grafica.Data = RParos(1)
                                        'SI ES EN MINUTOS
                                        If OptMinutos.Value = True Then
                                            VEtiqueta = (RParos(1))
                                        'SI ES HORAS
                                        Else
                                            VEtiqueta = Format((RParos(1)), "#,###,##0.00")
                                        End If
                                        'BUSCA LA DESCRIPCION DEL PARO
                                        Set RBuscaDescripcionParo = Db.OpenRecordset("Select DescripcionParo From Paros Where CodigoParo = '" & RParos(0) & "'")
                                            If RBuscaDescripcionParo.RecordCount > 0 Then
                                                'Grafica.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & Left(RBuscaDescripcionParo!DescripcionParo, 15)
                                                Grafica.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & RBuscaDescripcionParo!DescripcionParo
                                            Else
                                                Grafica.ColumnLabel = ""
                                            End If
                                        
                                        cont = cont + 1
                                    RParos.MoveNext
                            Loop
                                                        
'PARA LOS PAROS N______________________________________________________________________________________________________
            
                            cont = 1
                            CantidadDeColumnas = 0
                            
                            'CUENTA CUANDOS GRUPOS SE CREAN PARA CALCULAR LAS COLUMNAS DEL GRAFICA
                            Set RCuentaParos = Db.OpenRecordset("Select Top 10 Count(*) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And P.CodigoParo = DP.Paro And P.Tipo = 'N' Group By DP.Paro")
                            If RCuentaParos.RecordCount > 0 Then
                                Do Until RCuentaParos.EOF
                                            CantidadDeColumnas = CantidadDeColumnas + 1
                                        RCuentaParos.MoveNext
                                Loop
                            Else
                                CantidadDeColumnas = 0
                            End If

                            
                             'AGRUPA LOS PAROS POR CODIGO DE PARO Y TIPO 'N'
                            Set RParosN = Db.OpenRecordset("Select Top 10 Paro, Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' Group By DP.Paro")
                            
                            Do Until RParosN.EOF
                                    'AGREGA DATOS A LA BASE DE DATOS
                                    RGraficaDeParos.AddNew
                                            RGraficaDeParos!CodigoParo = RParosN(0)
                                            'MINUTOS
                                            If OptMinutos.Value = True Then
                                                RGraficaDeParos!Minutos = RParosN(1)
                                            'HORAS
                                            Else
                                                RGraficaDeParos!Minutos = Format((RParosN(1) / 60), "#,###,##0.00")
                                            End If
                                            RGraficaDeParos!Tipo = "N"
                                    RGraficaDeParos.Update
                                RParosN.MoveNext
                            Loop
                                    
                            'SACA TODOS LOS PAROS DE LA BASE DE DATOS ORDENADOS POR MINUTOS
                            Set RParosN = Db.OpenRecordset("Select * from GraficaDeParos Where Tipo = 'N' Order by Minutos")
                            
                            'CREA LA CANTIDAD DE COLUMNAS QUE VA A TENER LA GRAFICA
                            GraficaN.ColumnCount = CantidadDeColumnas
                                        
                            Do Until RParosN.EOF
                                        GraficaN.Column = cont
                                        GraficaN.Data = RParosN(1)
                                        
                                        'SI ES EN MINUTOS
                                        If OptMinutos.Value = True Then
                                            VEtiqueta = (RParosN(1))
                                        'SI ES HORAS
                                        Else
                                            VEtiqueta = Format((RParosN(1)), "#,###,##0.00")
                                        End If
                                        
                                        
                                        'BUSCA LA DESCRIPCION DEL PARO
                                        Set RBuscaDescripcionParo = Db.OpenRecordset("Select DescripcionParo From Paros Where CodigoParo = '" & RParosN(0) & "'")
                                            If RBuscaDescripcionParo.RecordCount > 0 Then
                                                'GraficaN.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & Left(RBuscaDescripcionParo!DescripcionParo, 15)
                                                GraficaN.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & RBuscaDescripcionParo!DescripcionParo
                                            Else
                                                GraficaN.ColumnLabel = ""
                                            End If
                                        
                                        cont = cont + 1
                                    RParosN.MoveNext
                            Loop
'PARA LOS PAROS S______________________________________________________________________________________________________
                            
                            cont = 1
                            CantidadDeColumnas = 0
                            
                            'CUENTA CUANDOS GRUPOS SE CREAN PARA CALCULAR LAS COLUMNAS DEL GRAFICA
                            Set RCuentaParos = Db.OpenRecordset("Select Top 10 Count(*) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And P.CodigoParo = DP.Paro And P.Tipo = 'S' Group By DP.Paro")
                            If RCuentaParos.RecordCount > 0 Then
                                Do Until RCuentaParos.EOF
                                            CantidadDeColumnas = CantidadDeColumnas + 1
                                        RCuentaParos.MoveNext
                                Loop
                            Else
                                CantidadDeColumnas = 0
                            End If

                            
                             'AGRUPA LOS PAROS POR CODIGO DE PARO Y TIPO 'S'
                            Set RParosS = Db.OpenRecordset("Select Top 10 Paro, Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' Group By DP.Paro")
                                                                                    
                            Do Until RParosS.EOF
                                    'AGREGA DATOS A LA BASE DE DATOS
                                    RGraficaDeParos.AddNew
                                            RGraficaDeParos!CodigoParo = RParosS(0)
                                             'MINUTOS
                                            If OptMinutos.Value = True Then
                                                RGraficaDeParos!Minutos = RParosS(1)
                                            'HORAS
                                            Else
                                                RGraficaDeParos!Minutos = Format((RParosS(1) / 60), "#,###,##0.00")
                                            End If
                                            RGraficaDeParos!Tipo = "S"
                                    RGraficaDeParos.Update
                                RParosS.MoveNext
                            Loop
                                    
                            'SACA TODOS LOS PAROS DE LA BASE DE DATOS ORDENADOS POR MINUTOS
                            Set RParosS = Db.OpenRecordset("Select * from GraficaDeParos Where Tipo = 'S' Order by Minutos")
                            
                            'CREA LA CANTIDAD DE COLUMNAS QUE VA A TENER LA GRAFICA
                            GraficaS.ColumnCount = CantidadDeColumnas
                                        
                            Do Until RParosS.EOF
                                        GraficaS.Column = cont
                                        GraficaS.Data = RParosS(1)
                                        'SI ES EN MINUTOS
                                        If OptMinutos.Value = True Then
                                            VEtiqueta = (RParosS(1))
                                        'SI ES HORAS
                                        Else
                                            VEtiqueta = Format((RParosS(1)), "#,###,##0.00")
                                        End If
                                                                                
                                        'BUSCA LA DESCRIPCION DEL PARO
                                        Set RBuscaDescripcionParo = Db.OpenRecordset("Select DescripcionParo From Paros Where CodigoParo = '" & RParosS(0) & "'")
                                            If RBuscaDescripcionParo.RecordCount > 0 Then
                                                'GraficaS.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & Left(RBuscaDescripcionParo!DescripcionParo, 15)
                                                GraficaS.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & RBuscaDescripcionParo!DescripcionParo
                                            Else
                                                GraficaS.ColumnLabel = ""
                                            End If
                                        
                                        cont = cont + 1
                                    RParosS.MoveNext
                            Loop
                            
                            
    'TOTALES ________________________________________________________________________________________
                            
                                   'SUMA TODOS LOS MINUTOS DE TODOS LOS PAROS
                                        'Set RParos = Db.OpenRecordset("Select Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento")
                                        Set RParos = Db.OpenRecordset("Select Sum(minutos) from GraficaDeParos Where Tipo = 'T'")
                                            If RParos.RecordCount > 0 Then
                                                'MINUTOS
                                                If OptMinutos.Value = True Then
                                                    Grafica.TitleText = "Paros Y Produccion De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Minutos: " & RParos(0) & "      Linea: " & LblLin.Caption
                                                'HORAS
                                                Else
                                                    Grafica.TitleText = "Paros Y Produccion De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Horas: " & Format(RParos(0), "#,###,##0.00") & "      Linea " & LblLin.Caption
                                                End If
                                            End If
                                            
                                  'AGRUPA LOS PAROS POR CODIGO DE PARO Y TIPO 'N'
                                        Set RParosN = Db.OpenRecordset("Select Sum(minutos) from GraficaDeParos Where Tipo = 'N'")
                                        'Set RParosN = Db.OpenRecordset("Select Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N'")
                                                'MINUTOS
                                                If OptMinutos.Value = True Then
                                                    GraficaN.TitleText = "Paros N De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Minutos: " & RParosN(0) & "      Linea: " & LblLin.Caption
                                                'HORAS
                                                Else
                                                    GraficaN.TitleText = "Paros N De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Horas: " & Format(RParosN(0), "#,###,##0.00") & "      Linea: " & LblLin.Caption
                                                End If

                                    'AGRUPA LOS PAROS POR CODIGO DE PARO Y TIPO 'S'
                                        'Set RParosS = Db.OpenRecordset("Select Sum(DP.minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos As EP, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And EP.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & Txtlin.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S'")
                                        Set RParosS = Db.OpenRecordset("Select Sum(minutos) from GraficaDeParos Where Tipo = 'S'")
                                                'MINUTOS
                                                If OptMinutos.Value = True Then
                                                    GraficaS.TitleText = "Paros S De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Minutos: " & RParosS(0) & "      Linea: " & LblLin.Caption
                                                'HORAS
                                                Else
                                                    GraficaS.TitleText = "Paros S De " & DtpFecIni.Value & " A " & DtpFecFin.Value & "      Total Horas: " & Format(RParosS(0), "#,###,##0.00") & "      Linea: " & LblLin.Caption
                                                End If
                            
MousePointer = 0
End Sub

Private Sub CmdGrabar_Click()

       
   CDDialogo.CancelError = True
   On Error GoTo ErrHandler
       
    CDDialogo.InitDir = App.Path
    CDDialogo.ShowSave
    
    If tabGrafica.Tab = 0 Then
            Grafica.EditCopy
    ElseIf tabGrafica.Tab = 1 Then
            GraficaS.EditCopy
    ElseIf tabGrafica.Tab = 2 Then
            GraficaN.EditCopy
    End If
            SavePicture Clipboard.GetData, CDDialogo.FileName
            MsgBox "La gráfica ha sido guardada ", vbInformation, "Guardar gráfica"
    
ErrHandler:
  'User pressed the Cancel button
  Exit Sub


End Sub

Private Sub CmdImprimirGrafica_Click()
    MousePointer = 11
            If tabGrafica.Tab = 0 Then
                    Grafica.EditCopy
            ElseIf tabGrafica.Tab = 1 Then
                    GraficaS.EditCopy
            ElseIf tabGrafica.Tab = 2 Then
                    GraficaN.EditCopy
            End If

        Printer.Orientation = 2
        Printer.PaintPicture Clipboard.GetData, 0, 0
        
        

        
        Printer.EndDoc
        
    MousePointer = 0
    
        MsgBox "Grafica Impresa", vbOKOnly + vbInformation, "Informacion"

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
            Txtlin.Text = DBGridBusqueda.Columns(0).Text
            Txtlin.SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                Txtlin.Text = DBGridBusqueda.Columns(0).Text
                Txtlin.SetFocus
                FrameBusqueda.Visible = False
            End If

End Sub

Private Sub Form_Load()
        DataBusqueda.Connect = GConnect
        DataBusqueda.DatabaseName = BasedeDatos

        
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date
        
        
End Sub

Private Sub Form_Resize()
            tabGrafica.Height = Me.ScaleHeight - 600
            tabGrafica.Width = Me.ScaleWidth - 100
            
            Grafica.Height = Me.ScaleHeight - 1700
            Grafica.Width = Me.ScaleWidth - 500
            GraficaS.Height = Me.ScaleHeight - 1700
            GraficaS.Width = Me.ScaleWidth - 500
            GraficaN.Height = Me.ScaleHeight - 1700
            GraficaN.Width = Me.ScaleWidth - 500
            
End Sub

Private Sub TxtBusqueda_Change()
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

Private Sub TxtLin_Change()
        Set RLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RLineas.RecordCount > 0 Then
                LblLin.Caption = RLineas!Descrip
            Else
                LblLin.Caption = ""
            End If
        
End Sub

Private Sub Txtlin_DblClick()
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus

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
                    DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub
