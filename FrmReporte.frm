VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FrmReporte 
   Caption         =   "Reportes"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FrmReporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5400
      Picture         =   "FrmReporte.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cambiar Impresora"
      Top             =   120
      Width           =   375
   End
   Begin CRVIEWER9LibCtl.CRViewer9 Visor 
      Height          =   8535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      lastProp        =   500
      _cx             =   20981
      _cy             =   15055
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Crystal As New CRAXDRT.Application
Private Reporte As New CRAXDRT.Report
Private CPProperty As CRAXDRT.ConnectionProperty  '(VE) 4/10/2004
Private BasedeDatos As CRAXDRT.Database
Private Tablas As CRAXDRT.DatabaseTables
Private tabla As CRAXDRT.DatabaseTable
Dim SubReport As CRAXDRT.Report
Dim Sections As CRAXDRT.Sections
Dim Section As CRAXDRT.Section
Dim RepObjs As CRAXDRT.ReportObjects
Dim SubReportObj As CRAXDRT.SubreportObject

Dim i As Integer
Dim Cont As Integer
Dim vtexto As String
Dim n As Integer
Dim j As Integer

Private Sub Command1_Click()
    Reporte.PrinterSetupEx Me.hWnd
    Reporte.PrintOut True, 1
End Sub


Private Sub Form_Load()
On Error Resume Next

    
    
    Screen.MousePointer = vbHourglass
                    'ASIGNA A LA VARIABLE REPORTE EL NOMBRE Y RUTA DEL REPORTE
                    Set Reporte = Crystal.OpenReport(GRutaDeReportes & "\" & GNombreReporte, 1)
                    If Err <> 0 Then
                        MsgBox "error al cargar reporte" & Err.Number & Err.Description
                        Err.Clear
                    End If
                    
                    'Reporte.DiscardSavedData
                    Reporte.ReportTitle = GTituloReporte
                    Reporte.ReportComments = GComentarioReporte
                                    
                    
                                         
                        If GOrigenDeDatos = "AmaproAccess" Then

                                    Set BasedeDatos = Reporte.Database

                                            Set Tablas = BasedeDatos.Tables
                                                Cont = 1
                                                    For Each tabla In Tablas

                                                            Set CPProperty = tabla.ConnectionProperties("Database Name")
                                                            CPProperty.Value = GRutaDeReportes & "\Metalenvases.mdb"
                                                            Set CPProperty = tabla.ConnectionProperties("Database Password")
                                                            CPProperty.Value = "metal"

                                                        Cont = Cont + 1
                                                        If Err <> 0 Then
                                                            MsgBox "error en Tablas 1 " & Err.Number & " " & Err.Description
                                                            Err.Clear
                                                        End If
                                                   Next tabla

                                                    Set Sections = Reporte.Sections
                                                            For n = 1 To Sections.Count
                                                              Set Section = Sections.Item(n)
                                                              Set RepObjs = Section.ReportObjects

                                                             For i = 1 To RepObjs.Count
                                                               If RepObjs.Item(i).Kind = crSubreportObject Then
                                                                  Set SubReportObj = RepObjs.Item(i)
                                                                  Set SubReport = SubReportObj.OpenSubreport

                                                                                SubReport.FormulaSyntax = 0

                                                                                Set BasedeDatos = SubReport.Database
                                                                                Set Tablas = BasedeDatos.Tables
                                                                                Cont = 1
                                                                                For Each tabla In Tablas

                                                                                        Set CPProperty = tabla.ConnectionProperties("Database Name")
                                                                                        CPProperty.Value = GRutaDeReportes & "\Metalenvases.mdb"
                                                                                        Set CPProperty = tabla.ConnectionProperties("Database Password")
                                                                                        CPProperty.Value = "metal"

                                                                                   Cont = Cont + 1
                                                                                       If Err <> 0 Then
                                                                                            MsgBox "error en Tablas 2 subreportes " & Err.Number & " " & Err.Description
                                                                                            Err.Clear
                                                                                            Exit Sub
                                                                                        End If
                                                                                Next tabla

                                                                End If
                                                              Next i
                                                    Next n



                            Else 'ORACLE

                                For i = 1 To Reporte.Database.Tables.Count
                                   Set tabla = Reporte.Database.Tables(i)
                                       tabla.ConnectionProperties("User ID").Value = "produc"
                                       tabla.ConnectionProperties("password").Value = "produccio"
                                       tabla.ConnectionProperties("provider").Value = "MSDAORA"
                                Next
                            End If
                                                                    
                    
                    'SELECCIONA LOS DATOS DEL REPORTE
                    Reporte.RecordSelectionFormula = GCriteriaReporte
                    'ASIGNA EL REPORTE AL Visor
                    Visor.ReportSource = Reporte
                    Visor.ViewReport
                    'Visor.Zoom (100)
                                
                    If Err <> 0 Then
                        MsgBox "err" & Err.Number & Err.Description
                        Err.Clear
                    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    Visor.Top = 0
    Visor.Left = 0
    Visor.Height = ScaleHeight
    Visor.Width = ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
     
     Set Reporte = Nothing
     Set Crystal = Nothing
     If Err <> 0 Then
     End If

End Sub

