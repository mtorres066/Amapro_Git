VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GraficaDefectos 
   Caption         =   "Grafica De Defectos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
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
      Left            =   2880
      TabIndex        =   14
      Top             =   4080
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
         Picture         =   "GraficaDefectos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "GraficaDefectos.frx":2072
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "GraficaDefectos.frx":208D
         TabIndex        =   22
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
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
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Por Grupo"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Por Linea"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   11280
      Picture         =   "GraficaDefectos.frx":2A67
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdGrabar 
      Height          =   375
      Left            =   8400
      Picture         =   "GraficaDefectos.frx":4AD9
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdCopiar 
      Height          =   375
      Left            =   9000
      Picture         =   "GraficaDefectos.frx":500B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copiar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdImprimirGrafica 
      Height          =   375
      Left            =   9600
      Picture         =   "GraficaDefectos.frx":553D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Grafica"
      Top             =   120
      Width           =   495
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
      ItemData        =   "GraficaDefectos.frx":5A6F
      Left            =   1080
      List            =   "GraficaDefectos.frx":5A97
      TabIndex        =   5
      Text            =   "2dBar"
      Top             =   120
      Width           =   2415
   End
   Begin TabDlg.SSTab TabGrafica 
      Height          =   7200
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   12700
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Defectos De Captura Produccion"
      TabPicture(0)   =   "GraficaDefectos.frx":5B0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GraficaDefCapPro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Defectos De Producion Liberada"
      TabPicture(1)   =   "GraficaDefectos.frx":5C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GraficaDefProLib"
      Tab(1).ControlCount=   1
      Begin MSChart20Lib.MSChart GraficaDefCapPro 
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "GraficaDefectos.frx":5DC0
         TabIndex        =   4
         Top             =   480
         Width           =   11655
      End
      Begin MSChart20Lib.MSChart GraficaDefProLib 
         Height          =   6975
         Left            =   -74880
         OleObjectBlob   =   "GraficaDefectos.frx":73EE
         TabIndex        =   9
         Top             =   360
         Width           =   11655
      End
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61669379
      CurrentDate     =   37153
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61669379
      CurrentDate     =   37153
   End
   Begin VB.CommandButton CmdGenerar 
      Height          =   375
      Left            =   10680
      Picture         =   "GraficaDefectos.frx":8A1C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Generar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CDDialogo 
      Left            =   3480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp;JPEG"
      DialogTitle     =   "Grabar Grafica"
      Filter          =   "Pictures (*.bmp)|*.bmp"
      FilterIndex     =   3
   End
   Begin VB.Label LblDes 
      BackColor       =   &H80000016&
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
      Left            =   5640
      TabIndex        =   29
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label LblLin 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   28
      Top             =   720
      Width           =   1095
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
      Left            =   4080
      TabIndex        =   12
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
      Left            =   6120
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "GraficaDefectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RDefectos As Recordset
Dim RBuscaDefectos As Recordset
Dim Cont As Integer
Dim VTotalDefectos As Long
Dim VGranTotalDefectos As Long
Dim RLineas As Recordset

Private Sub CboVerGra_Click()
If CboVerGra.ListIndex = 0 Then
            GraficaDefCapPro.chartType = VtChChartType2dArea
            GraficaDefProLib.chartType = VtChChartType2dArea
ElseIf CboVerGra.ListIndex = 1 Then
            GraficaDefCapPro.chartType = VtChChartType2dBar
            GraficaDefProLib.chartType = VtChChartType2dBar
ElseIf CboVerGra.ListIndex = 2 Then
            GraficaDefCapPro.chartType = VtChChartType2dCombination
            GraficaDefProLib.chartType = VtChChartType2dCombination
ElseIf CboVerGra.ListIndex = 3 Then
            GraficaDefCapPro.chartType = VtChChartType2dLine
            GraficaDefProLib.chartType = VtChChartType2dLine
ElseIf CboVerGra.ListIndex = 4 Then
            GraficaDefCapPro.chartType = VtChChartType2dPie
            GraficaDefProLib.chartType = VtChChartType2dPie
ElseIf CboVerGra.ListIndex = 5 Then
            GraficaDefCapPro.chartType = VtChChartType2dStep
            GraficaDefProLib.chartType = VtChChartType2dStep
ElseIf CboVerGra.ListIndex = 6 Then
            GraficaDefCapPro.chartType = VtChChartType2dXY
            GraficaDefProLib.chartType = VtChChartType2dXY
ElseIf CboVerGra.ListIndex = 7 Then
            GraficaDefCapPro.chartType = VtChChartType3dArea
            GraficaDefProLib.chartType = VtChChartType3dArea
ElseIf CboVerGra.ListIndex = 8 Then
            GraficaDefCapPro.chartType = VtChChartType3dBar
            GraficaDefProLib.chartType = VtChChartType3dBar
ElseIf CboVerGra.ListIndex = 9 Then
            GraficaDefCapPro.chartType = VtChChartType3dCombination
            GraficaDefProLib.chartType = VtChChartType3dCombination
ElseIf CboVerGra.ListIndex = 10 Then
            GraficaDefCapPro.chartType = VtChChartType3dLine
            GraficaDefProLib.chartType = VtChChartType3dLine
ElseIf CboVerGra.ListIndex = 11 Then
            GraficaDefCapPro.chartType = VtChChartType3dStep
            GraficaDefProLib.chartType = VtChChartType3dStep
End If

End Sub

Private Sub CmdCopiar_Click()
If TabGrafica.Tab = 0 Then
        GraficaDefCapPro.EditCopy
ElseIf TabGrafica.Tab = 1 Then
        GraficaDefProLib.EditCopy
End If

End Sub

Private Sub CmdGenerar_Click()
MousePointer = 11
            Set RDefectos = Db.OpenRecordset("Select Defecto, Descrip From Defectos")
            If RDefectos.RecordCount > 0 Then
                            Cont = 1
                            VTotalDefectos = 0
                            VGranTotalDefectos = 0
                            Do Until RDefectos.EOF
                                            'PONE CUANTAS COLUMNAS VA A TENER LA GRAFICA
                                            GraficaDefCapPro.ColumnCount = Cont
                                            
                                            'OPCION DE TODOS
                                            If OptOpcion.Item(0).Value = True Then
                                                'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                Set RBuscaDefectos = Db.OpenRecordset("Select Sum(Cantidad) From ProduccionConDefectos Where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Defecto = '" & RDefectos!Defecto & "'")
                                                    If IsNull(RBuscaDefectos(0)) Then
                                                            VTotalDefectos = 0
                                                    Else
                                                            VTotalDefectos = RBuscaDefectos(0)
                                                    End If
                                            'OPCION DE GRUPO
                                            ElseIf OptOpcion.Item(1).Value = True Then
                                                'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                Set RBuscaDefectos = Db.OpenRecordset("Select Sum(P.Cantidad) From ProduccionConDefectos As P, Lineas as L Where P.Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And P.Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Defecto = '" & RDefectos!Defecto & "' And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'")
                                                    If IsNull(RBuscaDefectos(0)) Then
                                                            VTotalDefectos = 0
                                                    Else
                                                            VTotalDefectos = RBuscaDefectos(0)
                                                    End If
                                            'OPCION DE LINEA
                                            ElseIf OptOpcion.Item(2).Value = True Then
                                                'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                Set RBuscaDefectos = Db.OpenRecordset("Select Sum(P.Cantidad) From ProduccionConDefectos As P, Lineas as L Where P.Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And P.Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Defecto = '" & RDefectos!Defecto & "' And P.Linea = L.Linea And L.Linea = '" & Txtlin.Text & "'")
                                                    If IsNull(RBuscaDefectos(0)) Then
                                                            VTotalDefectos = 0
                                                    Else
                                                            VTotalDefectos = RBuscaDefectos(0)
                                                    End If
                                            End If
                                                
                                            'SUMA EL TOTAL DE TODOS LOS DEFECTOS
                                            VGranTotalDefectos = VGranTotalDefectos + VTotalDefectos
                                            
                                            If VTotalDefectos > 0 Then
                                                        GraficaDefCapPro.Column = Cont
                                                        GraficaDefCapPro.Data = VTotalDefectos
                                                        GraficaDefCapPro.ColumnLabel = Left(RDefectos!Descrip & Space(40), 40) & Right(Space(7) & Format(VTotalDefectos, "#,###,###"), 7) & Space(2)
                                                        Cont = Cont + 1
                                            End If
                                        
                                    RDefectos.MoveNext
                            Loop
                                    If VTotalDefectos = 0 Then
                                            GraficaDefCapPro.ColumnCount = GraficaDefCapPro.ColumnCount - 1
                                    End If
                            
                                    GraficaDefCapPro.Title = "Total De Defectos " & Format(VGranTotalDefectos, "#,###,###")
                                    
            End If
            
            
            
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                    'CAPTURA DE PRODUCCION LIBERADA
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            Set RDefectos = Db.OpenRecordset("Select Defecto, Descrip From Defectos")
            If RDefectos.RecordCount > 0 Then
                            Cont = 1
                            VTotalDefectos = 0
                            VGranTotalDefectos = 0
                            Do Until RDefectos.EOF
                                            'PONE CUANTAS COLUMNAS VA A TENER LA GRAFICA
                                            GraficaDefProLib.ColumnCount = Cont
                                            
                                            'OPCION DE TODOS
                                            If OptOpcion.Item(0).Value = True Then
                                                    'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                    Set RBuscaDefectos = Db.OpenRecordset("Select Sum(Cantidad) From ProduccionLiberadaConDefectos Where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Defecto = '" & RDefectos!Defecto & "'")
                                                            If IsNull(RBuscaDefectos(0)) Then
                                                                    VTotalDefectos = 0
                                                            Else
                                                                    VTotalDefectos = RBuscaDefectos(0)
                                                            End If
                                            'OPCION DE GRUPO
                                            ElseIf OptOpcion.Item(1).Value = True Then
                                                'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                    Set RBuscaDefectos = Db.OpenRecordset("Select Sum(P.Cantidad) From ProduccionLiberadaConDefectos AS P, Lineas as L Where P.Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And P.Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Defecto = '" & RDefectos!Defecto & "' And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'")
                                                            If IsNull(RBuscaDefectos(0)) Then
                                                                    VTotalDefectos = 0
                                                            Else
                                                                    VTotalDefectos = RBuscaDefectos(0)
                                                            End If
                                            'OPCION DE LINEA
                                            ElseIf OptOpcion.Item(2).Value = True Then
                                                    'SUMA LA CANTIDAD DE DEFECTOS DENTRO DE LAS FECHAS
                                                    Set RBuscaDefectos = Db.OpenRecordset("Select Sum(P.Cantidad) From ProduccionLiberadaConDefectos AS P, Lineas as L Where P.Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And P.Fec_prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Defecto = '" & RDefectos!Defecto & "' And P.Linea = L.Linea And L.Linea = '" & Txtlin.Text & "'")
                                                            If IsNull(RBuscaDefectos(0)) Then
                                                                    VTotalDefectos = 0
                                                            Else
                                                                    VTotalDefectos = RBuscaDefectos(0)
                                                            End If
                                            End If
                                                    
                                            'SUMA EL TOTAL DE TODOS LOS DEFECTOS
                                            VGranTotalDefectos = VGranTotalDefectos + VTotalDefectos
                                            
                                            If VTotalDefectos > 0 Then
                                                        GraficaDefProLib.Column = Cont
                                                        GraficaDefProLib.Data = VTotalDefectos
                                                        GraficaDefProLib.ColumnLabel = Left(RDefectos!Descrip & Space(40), 40) & Right(Space(7) & Format(VTotalDefectos, "#,###,###"), 7) & Space(2)
                                                        Cont = Cont + 1
                                            End If
                                        
                                    RDefectos.MoveNext
                            Loop
                                    If VTotalDefectos = 0 Then
                                            GraficaDefProLib.ColumnCount = GraficaDefProLib.ColumnCount - 1
                                    End If
                            
                                    GraficaDefProLib.Title = "Total De Defectos " & Format(VGranTotalDefectos, "#,###,###")
                                    
            End If
            


MousePointer = 0
End Sub

Private Sub CmdGrabar_Click()

       
   CDDialogo.CancelError = True
   On Error GoTo ErrHandler
       
    CDDialogo.InitDir = App.Path
    CDDialogo.ShowSave
    
    If TabGrafica.Tab = 0 Then
            GraficaDefCapPro.EditCopy
    ElseIf TabGrafica.Tab = 1 Then
            GraficaDefProLib.EditCopy
    End If
             
            SavePicture Clipboard.GetData, CDDialogo.FileName
            MsgBox "La gráfica ha sido guardada ", vbInformation, "Guardar gráfica"
    
ErrHandler:
  'User pressed the Cancel button
  Exit Sub


End Sub

Private Sub CmdImprimirGrafica_Click()
    MousePointer = 11
    If TabGrafica.Tab = 0 Then
            GraficaDefCapPro.EditCopy
    ElseIf TabGrafica.Tab = 1 Then
            GraficaDefProLib.EditCopy
    End If

        Printer.Orientation = 2
        Printer.PaintPicture Clipboard.GetData, 0, 0
        
        Printer.EndDoc
    MousePointer = 0

End Sub

Private Sub Command1_Click()
        Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        'GRUPO
        If OptOpcion.Item(1).Value = True Then
            Txtlin.Text = DBGridBusqueda.Columns(2)
        'LINEA
        ElseIf OptOpcion.Item(2).Value = True Then
            Txtlin.Text = DBGridBusqueda.Columns(0)
        End If
        FrameBusqueda.Visible = False
        Txtlin.SetFocus

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                'GRUPO
                If OptOpcion.Item(1).Value = True Then
                    Txtlin.Text = DBGridBusqueda.Columns(2)
                'LINEA
                ElseIf OptOpcion.Item(2).Value = True Then
                    Txtlin.Text = DBGridBusqueda.Columns(0)
                End If
                FrameBusqueda.Visible = False
                Txtlin.SetFocus
        End If

End Sub

Private Sub Form_Load()
        DTPFecIni.Value = Date
        DTPFecFin.Value = Date

            DataBusqueda.ConnectionString = GTipoProveedor
            DataBusqueda.Refresh

End Sub

Private Sub Form_Resize()
                        
            TabGrafica.Height = Me.ScaleHeight - 600
            TabGrafica.Width = Me.ScaleWidth - 100
            
            GraficaDefCapPro.Height = Me.ScaleHeight - 1500
            GraficaDefCapPro.Width = Me.ScaleWidth - 400
            
            GraficaDefProLib.Height = Me.ScaleHeight - 1500
            GraficaDefProLib.Width = Me.ScaleWidth - 400
            
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        'TODOS
        If Index = 0 Then
            Txtlin.Visible = False
            LblLin.Caption = ""
            LblDes.Caption = ""
        'GRUPO
        ElseIf Index = 1 Then
            Txtlin.Visible = True
            LblLin.Caption = "Grupo"
            Txtlin.SetFocus
        'LINEA
        ElseIf Index = 2 Then
            Txtlin.Visible = True
            LblLin.Caption = "Linea"
            Txtlin.SetFocus
        End If

End Sub

Private Sub Txtbusqueda_Change()
            
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            'OPCION CUALQUIER PALABRA
                            If OptTipo.Item(1).Value = True Then
                                    DataBusqueda.RecordSource = ("Select Esp_Tec, Descrip, Size, Origen, Envases from FichaTecnica Where Descrip Like '*" & Txtbusqueda.Text & "*'")
                            'OPCION PALABRA INICIAL
                            ElseIf OptTipo.Item(0).Value = True Then
                                    DataBusqueda.RecordSource = ("Select Esp_Tec, Descrip, Size, Origen, Envases from FichaTecnica Where Descrip Like '" & Txtbusqueda.Text & "*'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION CUALQUIER PALABRA
                            If OptTipo.Item(1).Value = True Then
                                DataBusqueda.RecordSource = ("Select Esp_Tec, Descrip, Size, Origen, Envases from FichaTecnica Where Esp_Tec Descrip Like '*" & Txtbusqueda.Text & "*'")
                            'OPCION PALABRA INICIAL
                            ElseIf OptTipo.Item(0).Value = True Then
                                DataBusqueda.RecordSource = ("Select Esp_Tec, Descrip, Size, Origen, Envases from FichaTecnica Where Esp_Tec Like '" & Txtbusqueda.Text & "*'")
                            End If
                    End If
                            DataBusqueda.Refresh
                            DBGridBusqueda.Refresh
                            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtLin_Change()
        Set RLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RLineas.RecordCount > 0 Then
                LblDes.Caption = RLineas!Descrip
            End If

End Sub

Private Sub Txtlin_DblClick()
        FrameBusqueda.Visible = True
        DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
        DataBusqueda.Refresh
        DBGridBusqueda.Refresh
        DBGridBusqueda.Columns(1).Width = "4000"
        Txtbusqueda.SetFocus

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
            FrameBusqueda.Visible = True
            DataBusqueda.RecordSource = "Select Linea, Descrip, Grupo From Lineas"
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
            Txtbusqueda.SetFocus
        End If

End Sub
