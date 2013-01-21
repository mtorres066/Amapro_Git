VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GraficaDesperdicio 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Grafica De DESPERDICIO"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "GraficaDesperdicio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabGrafica 
      Height          =   7935
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Desperdicio En Proceso"
      TabPicture(0)   =   "GraficaDesperdicio.frx":548A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GraficaProceso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Desperdicio De Proveedor"
      TabPicture(1)   =   "GraficaDesperdicio.frx":5984
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GraficaProveedor"
      Tab(1).ControlCount=   1
      Begin MSChart20Lib.MSChart GraficaProceso 
         Height          =   6975
         Left            =   120
         OleObjectBlob   =   "GraficaDesperdicio.frx":625E
         TabIndex        =   12
         Top             =   720
         Width           =   11535
      End
      Begin MSChart20Lib.MSChart GraficaProveedor 
         Height          =   7095
         Left            =   -74880
         OleObjectBlob   =   "GraficaDesperdicio.frx":7C20
         TabIndex        =   13
         Top             =   720
         Width           =   11535
      End
   End
   Begin MSComDlg.CommonDialog CDDialogo 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp;JPEG"
      DialogTitle     =   "Grabar Grafica"
      Filter          =   "Pictures (*.bmp)|*.bmp"
      FilterIndex     =   3
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   10920
      Picture         =   "GraficaDesperdicio.frx":948E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salida"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton CmdGrabar 
      Height          =   495
      Left            =   7800
      Picture         =   "GraficaDesperdicio.frx":B500
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Grabar Grafica"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdCopiar 
      Height          =   495
      Left            =   8520
      Picture         =   "GraficaDesperdicio.frx":BA32
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copiar Grafica"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdImprimirGrafica 
      Height          =   495
      Left            =   9240
      Picture         =   "GraficaDesperdicio.frx":BF64
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Grafica"
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox CboVerGra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "GraficaDesperdicio.frx":C496
      Left            =   1080
      List            =   "GraficaDesperdicio.frx":C4C1
      TabIndex        =   3
      Text            =   "2dBar"
      Top             =   120
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24641539
      CurrentDate     =   37153
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24641539
      CurrentDate     =   37153
   End
   Begin VB.CommandButton CmdGenerar 
      Default         =   -1  'True
      Height          =   495
      Left            =   9960
      Picture         =   "GraficaDesperdicio.frx":C538
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Generar Grafica"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Vistas De Grafica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "GraficaDesperdicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RCuentaProcesos As Recordset
Dim RProcesos As Recordset
Dim RParosS As Recordset
Dim RParosN As Recordset
Dim RBuscaDescripcionProceso As Recordset

Dim CantidadDeColumnas As Integer
Dim Cont As Integer
Dim VEtiqueta As String


'Dim cnn As New ADODB.Connection
'Dim rst As New ADODB.Recordset
'Dim strProveedor As String
'Dim strOrigendatos As String
'Dim strSQL As String


Private Sub CboVerGra_Click()
If CboVerGra.ListIndex = 0 Then
            GraficaProceso.chartType = VtChChartType2dArea
            GraficaProveedor.chartType = VtChChartType2dArea
            
ElseIf CboVerGra.ListIndex = 1 Then
            GraficaProceso.chartType = VtChChartType2dBar
            GraficaProveedor.chartType = VtChChartType2dBar
            
ElseIf CboVerGra.ListIndex = 2 Then
            GraficaProceso.chartType = VtChChartType2dCombination
            GraficaProveedor.chartType = VtChChartType2dCombination
            
ElseIf CboVerGra.ListIndex = 3 Then
            GraficaProceso.chartType = VtChChartType2dLine
            GraficaProveedor.chartType = VtChChartType2dLine
            
ElseIf CboVerGra.ListIndex = 4 Then
            GraficaProceso.chartType = VtChChartType2dPie
            GraficaProveedor.chartType = VtChChartType2dPie
            
ElseIf CboVerGra.ListIndex = 5 Then
            GraficaProceso.chartType = VtChChartType2dStep
            GraficaProveedor.chartType = VtChChartType2dStep
            
ElseIf CboVerGra.ListIndex = 6 Then
            GraficaProceso.chartType = VtChChartType2dXY
            GraficaProveedor.chartType = VtChChartType2dXY
            
ElseIf CboVerGra.ListIndex = 7 Then
            GraficaProceso.chartType = VtChChartType3dArea
            GraficaProveedor.chartType = VtChChartType3dArea
            
ElseIf CboVerGra.ListIndex = 8 Then
            GraficaProceso.chartType = VtChChartType3dBar
            GraficaProveedor.chartType = VtChChartType3dBar
            
ElseIf CboVerGra.ListIndex = 9 Then
            GraficaProceso.chartType = VtChChartType3dCombination
            GraficaProveedor.chartType = VtChChartType3dCombination
            
ElseIf CboVerGra.ListIndex = 10 Then
            GraficaProceso.chartType = VtChChartType3dLine
            GraficaProveedor.chartType = VtChChartType3dLine
            
ElseIf CboVerGra.ListIndex = 11 Then
            GraficaProceso.chartType = VtChChartType3dStep
            GraficaProveedor.chartType = VtChChartType3dStep
            
End If

End Sub

Private Sub CmdCopiar_Click()
    If tabGrafica.Tab = 0 Then
            GraficaProceso.EditCopy
    ElseIf tabGrafica.Tab = 1 Then
            GraficaProveedor.EditCopy
    End If
        

End Sub

Private Sub CmdGenerar_Click()
MousePointer = 11
On Error Resume Next
            
            'strSQL = "Select paro, sum(minutos) From DetalleCapturaParos Group By Paro"
            'strProveedor = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & BasedeDatos & ";Mode=ReadWrite"
            'cnn.Open strProveedor
            'rst.Open strSQL, cnn, adOpenStatic
            'Set GraficaProceso.DataSource = rst
            'Set DataGrid1.DataSource = rst
            'If Err <> 0 Then
            '    MsgBox Err.Number & Err.Description
            'End If
            'rst.Close
            'cnn.Close
            
    'PROCESO ________________________________________________________________________________________________
            
            CantidadDeColumnas = 0
            Cont = 1
            
                        'CUENTA CUANDOS GRUPOS SE CREAN PARA CALCULAR LAS COLUMNAS DEL GraficaProceso
                        Set RCuentaProcesos = Db.OpenRecordset("Select Count(*) From CapturaDesperdicioMateriaPrima Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Group By CodigoProceso")
                            If RCuentaProcesos.RecordCount > 0 Then
                                Do Until RCuentaProcesos.EOF
                                            CantidadDeColumnas = CantidadDeColumnas + 1
                                        RCuentaProcesos.MoveNext
                                Loop
                            Else
                                CantidadDeColumnas = 0
                            End If
            
                            'AGRUPA LOS PROCESOS POR CODIGO DE PROCESO
                            Set RProcesos = Db.OpenRecordset("Select CodigoProceso, Sum(CuerposProceso) From CapturaDesperdicioMateriaPrima Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Group By CodigoProceso")
                            
                            'CREA LA CANTIDAD DE COLUMNAS QUE VA A TENER LA GraficaProceso
                            GraficaProceso.ColumnCount = CantidadDeColumnas
                                        
                            Do Until RProcesos.EOF
                                    'CUERPOS POR PROCESO
                                        GraficaProceso.Column = Cont
                                        GraficaProceso.Data = RProcesos(1)
                                        
                                        VEtiqueta = Str(RProcesos(1))
                                        'BUSCA LA DESCRIPCION DEL PROCESO
                                        Set RBuscaDescripcionProceso = Db.OpenRecordset("Select Descripcion From ProcesosMateriaPrima Where CodigoProceso = '" & RProcesos(0) & "'")
                                            If RBuscaDescripcionProceso.RecordCount > 0 Then
                                                GraficaProceso.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & RBuscaDescripcionProceso!Descripcion
                                            Else
                                                GraficaProceso.ColumnLabel = ""
                                            End If
                                            
                                        Cont = Cont + 1
                                    RProcesos.MoveNext
                            Loop
                            
                            
    'PROVEEDOR____________________________________________________________________________________________________
    
    
                        
                        Cont = 1
            
                        
                            'AGRUPA LOS PROCESOS POR CODIGO DE PROCESO
                            Set RProcesos = Db.OpenRecordset("Select CodigoProceso, Sum(CuerposProveedor) From CapturaDesperdicioMateriaPrima Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Group By CodigoProceso")
                            
                            'CREA LA CANTIDAD DE COLUMNAS QUE VA A TENER LA GRAFICA PROVEEDOR
                            GraficaProveedor.ColumnCount = CantidadDeColumnas
                                        
                            Do Until RProcesos.EOF
                                    'CUERPOS POR PROCESO
                                        GraficaProveedor.Column = Cont
                                        GraficaProveedor.Data = RProcesos(1)
                                        VEtiqueta = Str(RProcesos(1))
                                        'BUSCA LA DESCRIPCION DEL PROCESO
                                        Set RBuscaDescripcionProceso = Db.OpenRecordset("Select Descripcion From ProcesosMateriaPrima Where CodigoProceso = '" & RProcesos(0) & "'")
                                            If RBuscaDescripcionProceso.RecordCount > 0 Then
                                                'CONVIERTE LOS MINUTOS A HORAS
                                                GraficaProveedor.ColumnLabel = Left(VEtiqueta & Space(6), 6) & Space(2) & RBuscaDescripcionProceso!Descripcion
                                            Else
                                                GraficaProveedor.ColumnLabel = ""
                                            End If
                                        
                                        Cont = Cont + 1
                                    RProcesos.MoveNext
                            Loop
                            
                          

                            
MousePointer = 0
End Sub

Private Sub CmdGrabar_Click()

       
   CDDialogo.CancelError = True
   On Error GoTo ErrHandler
       
    CDDialogo.InitDir = App.Path
    CDDialogo.ShowSave
    
    If tabGrafica.Tab = 0 Then
            GraficaProceso.EditCopy
    ElseIf tabGrafica.Tab = 1 Then
            GraficaProveedor.EditCopy
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
                    GraficaProceso.EditCopy
            ElseIf tabGrafica.Tab = 1 Then
                    GraficaProveedor.EditCopy
            End If


        Printer.PaintPicture Clipboard.GetData, 0, 0
        
        Printer.EndDoc
    MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        DTPFecIni.Value = Date
        DTPFecFin.Value = Date
        
            
            
End Sub

Private Sub Form_Resize()
            tabGrafica.Height = Me.ScaleHeight - 600
            tabGrafica.Width = Me.ScaleWidth - 100
            GraficaProceso.Height = Me.ScaleHeight - 1700
            GraficaProceso.Width = Me.ScaleWidth - 500
            GraficaProveedor.Height = Me.ScaleHeight - 1700
            GraficaProveedor.Width = Me.ScaleWidth - 500
            

End Sub
