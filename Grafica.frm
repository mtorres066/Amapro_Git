VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Graficas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graficas De Produccion Terminada"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "Grafica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
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
         TabIndex        =   26
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "Grafica.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Bindings        =   "Grafica.frx":24B4
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "Grafica.frx":24CF
         TabIndex        =   27
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
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
      Left            =   9360
      TabIndex        =   21
      Top             =   8280
      Width           =   2415
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   18
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
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Por Linea"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Por Grupo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo De Calidad"
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
      Height          =   975
      Left            =   1920
      TabIndex        =   9
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton OptCalidad 
         Caption         =   "Producto No Conforme"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OptCalidad 
         Caption         =   "Producto Conforme"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton OptCalidad 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdImprimirGrafica 
      Height          =   375
      Left            =   9480
      Picture         =   "Grafica.frx":2EA9
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdCopiar 
      Height          =   375
      Left            =   8880
      Picture         =   "Grafica.frx":33DB
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copiar Grafica"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdGrabar 
      Height          =   375
      Left            =   8280
      Picture         =   "Grafica.frx":390D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Grabar Grafica"
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
      ItemData        =   "Grafica.frx":3E3F
      Left            =   4200
      List            =   "Grafica.frx":3E67
      TabIndex        =   4
      Text            =   "2dBar"
      Top             =   240
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DtPickerAño 
      Height          =   315
      Left            =   7200
      TabIndex        =   2
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   23986179
      UpDown          =   -1  'True
      CurrentDate     =   36870
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11280
      Picture         =   "Grafica.frx":3EDC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sale de Graficas"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdGeneraGrafica 
      Height          =   375
      Left            =   10680
      Picture         =   "Grafica.frx":5F4E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Genera Grafica"
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CDDialogo 
      Left            =   10080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp;JPEG"
      DialogTitle     =   "Grabar Grafica"
      Filter          =   "Pictures (*.bmp)|*.bmp"
      FilterIndex     =   3
   End
   Begin MSChart20Lib.MSChart Grafica 
      Height          =   7215
      Left            =   0
      OleObjectBlob   =   "Grafica.frx":7FC0
      TabIndex        =   17
      Top             =   1080
      Width           =   11775
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
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Width           =   1215
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
      Left            =   6360
      TabIndex        =   19
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Graficas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RTotal As Recordset
Dim RTotales As Recordset
Dim RLineas As Recordset
Dim Columnas As String
Dim Tablas As String
Dim Criteria As String
Dim VAño As Double
Dim VMes As String
Dim VTotal As Long
Dim Cont As Integer


Private Sub CmdCopiar_Click()
        Grafica.EditCopy
End Sub


Private Sub CboVerGra_Click()
If CboVerGra.ListIndex = 0 Then
            Grafica.chartType = VtChChartType2dArea
            
ElseIf CboVerGra.ListIndex = 1 Then
            Grafica.chartType = VtChChartType2dBar
ElseIf CboVerGra.ListIndex = 2 Then
            Grafica.chartType = VtChChartType2dCombination
            
ElseIf CboVerGra.ListIndex = 3 Then
            Grafica.chartType = VtChChartType2dLine
            
ElseIf CboVerGra.ListIndex = 4 Then
            Grafica.chartType = VtChChartType2dPie
ElseIf CboVerGra.ListIndex = 5 Then
            Grafica.chartType = VtChChartType2dStep
            
ElseIf CboVerGra.ListIndex = 6 Then
            Grafica.chartType = VtChChartType2dXY
            
ElseIf CboVerGra.ListIndex = 7 Then
            Grafica.chartType = VtChChartType3dArea
ElseIf CboVerGra.ListIndex = 8 Then
            Grafica.chartType = VtChChartType3dBar

ElseIf CboVerGra.ListIndex = 9 Then
            Grafica.chartType = VtChChartType3dCombination
            
ElseIf CboVerGra.ListIndex = 10 Then
            Grafica.chartType = VtChChartType3dLine
            
ElseIf CboVerGra.ListIndex = 11 Then
            Grafica.chartType = VtChChartType3dStep
End If

End Sub

Private Sub CmdGeneraGrafica_Click()
MousePointer = 11
Cont = 1
VAño = Year(DtPickerAño.Value)
                        'CALIDAD TODAS_____________________________________________________________________________
                        If OptCalidad.Item(0).Value = True Then
                            'TODOS
                            If OptOpcion.Item(0).Value = True Then
                                        'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                        Set RTotales = Db.OpenRecordset("Select Count(*), Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño)
                                            If RTotales.RecordCount > 0 Then
                                                If IsNull(RTotales(0)) Then
                                                        VTotal = 0
                                                Else
                                                        VTotal = RTotales(1)
                                                        
                                                End If
                                            Else
                                                        VTotal = 0
                                            End If
                            'POR GRUPO
                            ElseIf OptOpcion.Item(1).Value = True Then
                              'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                        Set RTotales = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'")
                                            If RTotales.RecordCount > 0 Then
                                                If IsNull(RTotales(0)) Then
                                                        VTotal = 0
                                                Else
                                                        VTotal = RTotales(0)
                                                        
                                                End If
                                            Else
                                                        VTotal = 0
                                            End If

                            'POR LINEA
                            ElseIf OptOpcion.Item(2).Value = True Then
                                        'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                        Set RTotales = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " And Linea = '" & Txtlin.Text & "'")
                                            If RTotales.RecordCount > 0 Then
                                                If IsNull(RTotales(0)) Then
                                                        VTotal = 0
                                                Else
                                                        VTotal = RTotales(0)
                                                        
                                                End If
                                            Else
                                                        VTotal = 0
                                            End If

                            End If
                            
                        'PRODUCTO CONFORME _____________________________________________________________________________
                        ElseIf OptCalidad.Item(1).Value = True Then
                                    'TODOS
                                    If OptOpcion.Item(0).Value = True Then
                                                'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Count(*), Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(1)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
                                    'POR GRUPO
                                    ElseIf OptOpcion.Item(1).Value = True Then
                                      'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'" & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(0)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
        
                                    'POR LINEA
                                    ElseIf OptOpcion.Item(2).Value = True Then
                                                'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " And Linea = '" & Txtlin.Text & "'" & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(0)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
        
                                    End If
                        'PRODUCTO NO CONFORME_____________________________________________________________________________
                        ElseIf OptCalidad.Item(2).Value = True Then
                                    'TODOS
                                    If OptOpcion.Item(0).Value = True Then
                                                'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Count(*), Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " And Calidad = 'R'")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(1)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
                                    'POR GRUPO
                                    ElseIf OptOpcion.Item(1).Value = True Then
                                      'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'" & " And Calidad = 'R'")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(0)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
        
                                    'POR LINEA
                                    ElseIf OptOpcion.Item(2).Value = True Then
                                                'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                Set RTotales = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " And Linea = '" & Txtlin.Text & "'" & " And Calidad = 'R'")
                                                    If RTotales.RecordCount > 0 Then
                                                        If IsNull(RTotales(0)) Then
                                                                VTotal = 0
                                                        Else
                                                                VTotal = RTotales(0)
                                                                
                                                        End If
                                                    Else
                                                                VTotal = 0
                                                    End If
        
                                    End If
                        End If
                                  'DESPLIEGA EL TOTAL
                                  TxtTotal.Text = "Total:  " & Format(VTotal, "#,###,##0")
                        
                                  'SE CREA UN CONTADOR POR LOS 12 MESES
                                  Do Until Cont = 13
                                                If Cont = 1 Then
                                                    VMes = "Enero"
                                                ElseIf Cont = 2 Then
                                                    VMes = "Febrero"
                                                ElseIf Cont = 3 Then
                                                    VMes = "Marzo"
                                                ElseIf Cont = 4 Then
                                                    VMes = "Abril"
                                                ElseIf Cont = 5 Then
                                                    VMes = "Mayo"
                                                ElseIf Cont = 6 Then
                                                    VMes = "Junio"
                                                ElseIf Cont = 7 Then
                                                    VMes = "Julio"
                                                ElseIf Cont = 8 Then
                                                    VMes = "Agosto"
                                                ElseIf Cont = 9 Then
                                                    VMes = "Septiembre"
                                                ElseIf Cont = 10 Then
                                                    VMes = "Octubre"
                                                ElseIf Cont = 11 Then
                                                    VMes = "Noviembre"
                                                ElseIf Cont = 12 Then
                                                    VMes = "Diciembre"
                                                End If
                                                
                                                                                             
                                                                        'CALIDAD TODAS_____________________________________________________________________________
                                                            If OptCalidad.Item(0).Value = True Then
                                                                'TODOS
                                                                If OptOpcion.Item(0).Value = True Then
                                                                            'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                            Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont)
                                                                                         Grafica.Column = Cont
                                                                                         Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                                                                
                                                                                
                                                                'POR GRUPO
                                                                ElseIf OptOpcion.Item(1).Value = True Then
                                                                  'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                            Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If

                                    
                                                                'POR LINEA
                                                                ElseIf OptOpcion.Item(2).Value = True Then
                                                                            'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                            Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And Linea = '" & Txtlin.Text & "'")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                    
                                                                End If
                                                                
                                                            'PRODUCTO CONFORME _____________________________________________________________________________
                                                            ElseIf OptCalidad.Item(1).Value = True Then
                                                                        'TODOS
                                                                        If OptOpcion.Item(0).Value = True Then
                                                                                    'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                                                                        
                                                                        'POR GRUPO
                                                                        ElseIf OptOpcion.Item(1).Value = True Then
                                                                          'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'" & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                            
                                                                        'POR LINEA
                                                                        ElseIf OptOpcion.Item(2).Value = True Then
                                                                                    'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And Linea = '" & Txtlin.Text & "'" & " And (Calidad = 'A' or Calidad = 'I' Or Calidad = 'C')")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                            
                                                                        End If
                                                            'PRODUCTO NO CONFORME_____________________________________________________________________________
                                                            ElseIf OptCalidad.Item(2).Value = True Then
                                                                        'TODOS
                                                                        If OptOpcion.Item(0).Value = True Then
                                                                                    'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And Calidad = 'R'")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                                                        'POR GRUPO
                                                                        ElseIf OptOpcion.Item(1).Value = True Then
                                                                          'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) from Produccion AS P, Lineas as L where year(P.Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And P.Linea = L.Linea And L.Grupo = '" & Txtlin.Text & "'" & " And Calidad = 'R'")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                                                        'POR LINEA
                                                                        ElseIf OptOpcion.Item(2).Value = True Then
                                                                                    'TARIMAS Y ENVASES PRODUCTO CONFORME POR AÑO
                                                                                    Set RTotal = Db.OpenRecordset("Select Sum(Envases) from Produccion where year(Fec_Prd) = " & VAño & " and month(Fec_Prd) = " & Cont & " And Linea = '" & Txtlin.Text & "'" & " And Calidad = 'R'")
                                                                                        Grafica.Column = Cont
                                                                                        Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotal(0), "#,###,###"), 10) & Space(2)
                                                                                        
                                                                                        If RTotal(0) = 0 Then
                                                                                                  Grafica.Data = 0
                                                                                                  Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
                                                                                        Else
                                                                                                  Grafica.Data = RTotal(0) & Space(2)
                                                                                        End If
                                            
                                                                        End If
                                                            End If
                        
                                                                
                                                                
                                   Cont = Cont + 1
                                Loop
                  
                
    
 MousePointer = 0
    
End Sub

Private Sub CmdGrabar_Click()
       
   CDDialogo.CancelError = True
   On Error GoTo ErrHandler
       
    CDDialogo.InitDir = App.Path
    CDDialogo.ShowSave
    
    
            Grafica.EditCopy
             
            SavePicture Clipboard.GetData, CDDialogo.FileName
            MsgBox "La gráfica ha sido guardada ", vbInformation, "Guardar gráfica"
    
ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdImprimirGrafica_Click()
On Error Resume Next
    MousePointer = 11
    
    
                Grafica.EditCopy
            
        Printer.PaintPicture Clipboard.GetData, 0, 0
        
        If Err <> 0 Then
            MsgBox Err.Number & " " & Err.Description
            MousePointer = 0
            Exit Sub
        End If
        
        Cont = 1
        Do Until Cont = 40
                Printer.Print
            Cont = Cont + 1
        Loop
        
        
        Printer.FontSize = "14"
        Printer.FontBold = True
                        
    
                Printer.Print Space(85) & "Total: " & VTotal
                Printer.Print
                'CALIDAD TODAS
                If OptCalidad.Item(0).Value = True Then
                        'TODOS
                        If OptOpcion.Item(0).Value = True Then
                            Printer.Print Space(15) & "Grafica De Calidad Todas";
                        'GRUPO
                        ElseIf OptOpcion.Item(1).Value = True Then
                            Printer.Print Space(15) & "Grafica De Calidad Todas y Grupo " & Txtlin.Text;
                        'LINEA
                        ElseIf OptOpcion.Item(2).Value = True Then
                            Printer.Print Space(15) & "Grafica De Calidad Todas y Linea " & Txtlin.Text & " " & LblDes.Caption;
                        End If
                'PRODUCTO CONFORME
                ElseIf OptCalidad.Item(1).Value = True Then
                        'TODOS
                        If OptOpcion.Item(0).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto Conforme";
                        'GRUPO
                        ElseIf OptOpcion.Item(1).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto Conforme Y Grupo " & Txtlin.Text;
                        'LINEA
                        ElseIf OptOpcion.Item(2).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto Conforme Y Linea " & Txtlin.Text & " " & LblDes.Caption;
                        End If
                'PRODUCTO NO CONFORME
                ElseIf OptCalidad.Item(2).Value = True Then
                        'TODOS
                        If OptOpcion.Item(0).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto No Conforme";
                        'GRUPO
                        ElseIf OptOpcion.Item(1).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto No Conforme Y Grupo " & Txtlin.Text;
                        'LINEA
                        ElseIf OptOpcion.Item(2).Value = True Then
                            Printer.Print Space(15) & "Grafica De Producto No Conforme Y Linea " & Txtlin.Text & " " & LblDes.Caption;
                        End If
                End If
                
                
                Printer.Print " Del Año " & Format(DtPickerAño.Value, "yyyy")
        
        Printer.EndDoc
        
    MousePointer = 0
    
    MsgBox "Grafica Impresa"

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
            DtPickerAño.Value = Format(Date, "dd/mm/yyyy")
            DataBusqueda.ConnectionString = GTipoProveedor
            DataBusqueda.Refresh
    
End Sub

Private Sub Form_Resize()
'        TabGrafica.Width = Me.ScaleWidth - 200
'        TabGrafica.Height = Me.ScaleHeight - 900
'
'        Grafica.Width = Me.ScaleWidth - 500
'        Grafica.Height = Me.ScaleHeight - 1800
'        TxtTotTarPC.Move Me.ScaleWidth - 2500, Me.ScaleHeight - 1500
         
'        GraficaTarimasProductoNoConforme.Width = Me.ScaleWidth - 500
'        GraficaTarimasProductoNoConforme.Height = Me.ScaleHeight - 1800
'        TxtTotTarPNC.Move Me.ScaleWidth - 2500, Me.ScaleHeight - 1500
        
'        GraficaEnvases.Width = Me.ScaleWidth - 500
'        GraficaEnvases.Height = Me.ScaleHeight - 1800
'        TxtTotEnvPC.Move Me.ScaleWidth - 2500, Me.ScaleHeight - 1500
        
'        GraficaProductoNoConforme.Width = Me.ScaleWidth - 500
'        GraficaProductoNoConforme.Height = Me.ScaleHeight - 1800
'        TxtTotEnvPNC.Move Me.ScaleWidth - 2500, Me.ScaleHeight - 1500
        
 '       GraficaProductoNoConformeLiberado.Width = Me.ScaleWidth - 500
'        GraficaProductoNoConformeLiberado.Height = Me.ScaleHeight - 1800
'        TxtTotProNoConformeLiberado.Move Me.ScaleWidth - 2500, Me.ScaleHeight - 1500
               
        
        
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
                                    DataBusqueda.RecordSource = ("Select Linea, Descrip, Grupo from Lineas Where Descrip Like '*" & Txtbusqueda.Text & "*'")
                            'OPCION PALABRA INICIAL
                            ElseIf OptTipo.Item(0).Value = True Then
                                    DataBusqueda.RecordSource = ("Select Linea, Descrip, Grupo from Lineas Where Descrip Like '" & Txtbusqueda.Text & "*'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION CUALQUIER PALABRA
                            If OptTipo.Item(1).Value = True Then
                                DataBusqueda.RecordSource = ("Select Linea, Descrip, Grupo from Lineas Where Linea Like '*" & Txtbusqueda.Text & "*'")
                            'OPCION PALABRA INICIAL
                            ElseIf OptTipo.Item(0).Value = True Then
                                DataBusqueda.RecordSource = ("Select Linea, Descrip, Grupo from Lineas Where Linea Like '" & Txtbusqueda.Text & "*'")
                            End If
                    End If
                            DataBusqueda.Refresh
                            DBGridBusqueda.Refresh
                            DBGridBusqueda.Columns(1).Width = "4000"
                            
End Sub

Private Sub TxtLin_Change()
    If OptOpcion.Item(2).Value = True Then
        Set RLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RLineas.RecordCount > 0 Then
                LblDes.Caption = RLineas!Descrip
            Else
                LblDes.Caption = ""
            End If
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
