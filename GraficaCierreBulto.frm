VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GraficaCierreBulto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grafica De Cierres De Bulto (materia prima procesada)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "GraficaCierreBulto.frx":0000
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
      Height          =   8535
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   10680
         Picture         =   "GraficaCierreBulto.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
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
         Height          =   735
         Left            =   6360
         TabIndex        =   21
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "GraficaCierreBulto.frx":074C
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "GraficaCierreBulto.frx":0767
         TabIndex        =   27
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Label LblBusqueda 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
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
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
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
      Left            =   9240
      TabIndex        =   19
      Text            =   "Gran Total"
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   10440
      TabIndex        =   16
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame FrameTipoGrafica 
      Caption         =   "Tipo Consulta"
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
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton OptTipo 
         Caption         =   "Tipo De Materia Prima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo Materia Prima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton CmdImprimirGrafica 
      Height          =   615
      Left            =   9360
      Picture         =   "GraficaCierreBulto.frx":1141
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Grafica"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdCopiar 
      Height          =   615
      Left            =   8520
      Picture         =   "GraficaCierreBulto.frx":1673
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Copiar Grafica"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdGrabar 
      Height          =   615
      Left            =   7680
      Picture         =   "GraficaCierreBulto.frx":1BA5
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar Grafica"
      Top             =   120
      Width           =   735
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
      ItemData        =   "GraficaCierreBulto.frx":20D7
      Left            =   120
      List            =   "GraficaCierreBulto.frx":20FF
      TabIndex        =   3
      Text            =   "2dBar"
      Top             =   360
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DtPickerAño 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   360
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
      Format          =   62390275
      UpDown          =   -1  'True
      CurrentDate     =   36870
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   11040
      Picture         =   "GraficaCierreBulto.frx":2174
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sale de Graficas"
      Top             =   120
      Width           =   765
   End
   Begin VB.CommandButton CmdGeneraGrafica 
      Default         =   -1  'True
      Height          =   615
      Left            =   10200
      Picture         =   "GraficaCierreBulto.frx":41E6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Genera Grafica"
      Top             =   120
      Width           =   765
   End
   Begin MSComDlg.CommonDialog CDDialogo 
      Left            =   4200
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
      Height          =   7335
      Left            =   120
      OleObjectBlob   =   "GraficaCierreBulto.frx":6258
      TabIndex        =   15
      Top             =   1200
      Width           =   11655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "U/M"
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
      Left            =   9705
      TabIndex        =   18
      Top             =   840
      Width           =   390
   End
   Begin VB.Label LblUnidadMedida 
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
      Left            =   10200
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label LblDescripcion 
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
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label LblEtiqueta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Materia Prima"
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
      Left            =   330
      TabIndex        =   17
      Top             =   840
      Width           =   1605
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
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
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "GraficaCierreBulto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RTotales As Recordset
Dim RBuscaMateriaPrima As Recordset

Dim Columnas As String
Dim Tablas As String
Dim Criteria As String
Dim VAño As Double
Dim VMes As String
                   
Dim Cont As Integer

Dim BTipoMateriaPrima As Boolean
Dim BMateriaPrima As Boolean


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

            'SUMA TODO POR AÑO
            If OptCodigo.Value = True Then
                'SUMA EL TOTAL DESCARGADO EN EL MES Y DEPENDIENDO EL CODIGO DE MATERIA PRIMA
                Set RTotales = Db.OpenRecordset("Select sum(Total) from NumerosIngresosProcesados where year(Fecha) = " & VAño & " And CodigoMateriaPrima = '" & Txttexto.Text & "'")
            Else
                'SUMA EL TOTAL DESCARGADO EN EL MES Y DEPENDIENDO EL TIPO DE MATERIA PRIMA
                Set RTotales = Db.OpenRecordset("Select sum(N.Total) from NumerosIngresosProcesados AS N, CorrelativosMateriaPrima as CM where year(N.Fecha) = " & VAño & " And CM.TipoDeMateriaPrima = '" & Txttexto.Text & "' And CM.CodigoMateriaPrima = N.CodigoMateriaPrima")
            End If
            
                If RTotales.RecordCount > 0 Then
                    If IsNull(RTotales(0)) Then
                        TxtTotal.Text = "0"
                    Else
                        TxtTotal.Text = Format(RTotales(0), "#,###,##0")
                    End If
                End If
            

          
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
                                                
            'POR CODIGO
            If OptCodigo.Value = True Then
                'SUMA EL TOTAL DESCARGADO EN EL MES Y DEPENDIENDO EL CODIGO DE MATERIA PRIMA
                Set RTotales = Db.OpenRecordset("Select sum(Total) from NumerosIngresosProcesados where year(Fecha) = " & VAño & " And month(fecha) = " & Cont & " And CodigoMateriaPrima = '" & Txttexto.Text & "'")
            Else
                'SUMA EL TOTAL DESCARGADO EN EL MES Y DEPENDIENDO EL TIPO DE MATERIA PRIMA
                Set RTotales = Db.OpenRecordset("Select sum(N.Total) from NumerosIngresosProcesados AS N, CorrelativosMateriaPrima as CM where year(N.Fecha) = " & VAño & " And month(N.Fecha) = " & Cont & " And CM.TipoDeMateriaPrima = '" & Txttexto.Text & "' And CM.CodigoMateriaPrima = N.CodigoMateriaPrima")
            End If
            
            
            
            Grafica.Column = Cont
            
            If IsNull(RTotales(0)) Then
                Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
            Else
                Grafica.ColumnLabel = Left(VMes & Space(12), 12) & " " & Right(Space(10) & Format(RTotales(0), "#,###,###"), 10) & Space(2)
            End If
                                                                           
            If RTotales(0) = 0 Then
                Grafica.Data = 0
                Grafica.ColumnLabel = Left(VMes & Space(12), 12) & Space(10) & "0" & Space(2)
            Else
                If IsNull(RTotales(0)) Then
                    Grafica.Data = 0
                Else
                    Grafica.Data = RTotales(0) & Space(2)
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

Private Sub CmdImprimir_Click()
    MousePointer = 11
        Graficas.PrintForm
    MousePointer = 0
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
        
                Printer.Print Space(85) & "Total: " & TxtTotal.Text
                Printer.Print
                Printer.Print Space(15) & "Grafica De Materia Prima Procesada ";
                Printer.Print " Del Año " & Format(DtPickerAño.Value, "yyyy")
        
        Printer.EndDoc
        
    MousePointer = 0
    
    MsgBox "Grafica Impresa"

End Sub

Private Sub DBGridBusqueda_DblClick()
        Txttexto.Text = DBGridBusqueda.Columns(0)
        Txttexto.SetFocus
        FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            Txttexto.Text = DBGridBusqueda.Columns(0)
            Txttexto.SetFocus
            FrameBusqueda.Visible = False
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

Private Sub TabGrafica_Click(PreviousTab As Integer)
'        If TabGrafica.Tab = 4 Then
'            TxtTotProNoConformeLiberado.Visible = True
'        Else
'            TxtTotProNoConformeLiberado.Visible = False
'        End If
        
End Sub

Private Sub OptCodigo_Click()
        Txttexto.SetFocus
        Lbletiqueta.Caption = "Codigo Materia Prima"
End Sub

Private Sub OptTipo_Click()
        Txttexto.SetFocus
        Lbletiqueta.Caption = "Tipo Materia Prima"
End Sub

Private Sub Txtbusqueda_Change()
            
        'TIPOS DE MATERIA PRIMA EN CARPETA DE INVENTARIO
        If BTipoMateriaPrima = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where CodigoTipo Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where CodigoTipo Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        'CODIGO DE MATERIA PRIMA EN CARPETA DE INVENTARIO, TRASLADOS, ENTRADAS
        ElseIf BMateriaPrima = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        End If
End Sub

Private Sub TxtTexto_Change()
        'BUSCA LA DESCRIPCION DE LA MATERIA PRIMA DE ACUERDO A LA OPCION QUE SE ELIGE
        If OptCodigo.Value = True Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & Txttexto.Text & "'")
        Else
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select TMP.Descripcion, CMP.UnidadMedida From TiposDeMateriaPrima as TMP, CorrelativosMateriaPrima as CMP Where TMP.CodigoTipo = '" & Txttexto.Text & "' And TMP.CodigoTipo = CMP.TipoDeMateriaPrima")
        End If
            If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblDescripcion.Caption = RBuscaMateriaPrima!Descripcion
                    LblUnidadMedida.Caption = RBuscaMateriaPrima!UnidadMedida
            Else
                    LblDescripcion.Caption = ""
                    LblUnidadMedida.Caption = ""
            End If
        
End Sub

Private Sub TxtTexto_DblClick()
        If OptTipo.Value = True Then
                DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        Else
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        End If

End Sub

Private Sub TxtTexto_GotFocus()
        Txttexto.SelStart = 0
        Txttexto.SelLength = Len(Txttexto.Text)
End Sub

Private Sub TxtTexto_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 43 Then
        If OptTipo.Value = True Then
                DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        Else
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        End If
    End If

End Sub
