VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraCapturaRutinasTapas 
   Caption         =   "Genera Captura de Rutinas AUTOMATICA TAPAS"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   Icon            =   "GeneraCapturaRutinasTapas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DbGridSeametal 
      Height          =   7095
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12515
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
   Begin VB.Frame Frame1 
      Caption         =   "Tipo De Envase"
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
      Left            =   3360
      TabIndex        =   16
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton OptFondo 
         Caption         =   "Fondo"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptAnillo 
         Caption         =   "Anillo"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDos 
         Caption         =   "2 Piezas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Opttres 
         Caption         =   "3 Piezas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox TxtDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   7320
      Width           =   855
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtHor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox TxtFicTec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   7320
      Width           =   615
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
      Height          =   735
      Left            =   6720
      Picture         =   "GeneraCapturaRutinasTapas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton CmdCapturar 
      Caption         =   "&Capturar Rutinas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Picture         =   "GeneraCapturaRutinasTapas.frx":25FC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "Doc."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Hora"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Ficha Tecnica"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Linea"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   7200
      Width           =   8295
   End
End
Attribute VB_Name = "GeneraCapturaRutinasTapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaRutina As New ADODB.Recordset
Dim RMaximo As New ADODB.Recordset
Dim RDatos As New ADODB.Recordset
Dim RDatosdeRutina As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RSeaMetal As New ADODB.Recordset


Private Sub CmdCapturar_Click()
On Error Resume Next
MousePointer = 11
    
    'VERIFICA SI LA LINEA EXISTE
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where Linea = '" & TxtLin.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
        End If
        If RBuscaLinea.RecordCount > 0 Then
        Else
            MsgBox "Linea No Existe ", vbOKOnly + vbInformation, "Informacion"
            TxtLin.SetFocus
            Exit Sub
        End If
        
    'VALIDA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "informacion"
        MousePointer = 0
        Exit Sub
    End If

    'BUSCA LAS RUTINAS
    Set RBuscaRutina = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaRutina, "Select * from CapturaRutinas Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaRutina, "Select * from CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "'")
        End If
    If RBuscaRutina.RecordCount > 0 Then
    Else
        MsgBox "Rutina No Existe ", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
    End If
    
        
    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE QUE SON LAS MEDICIONES DEL SEAMETAL
    Set RDatos = New ADODB.Recordset
    Call Abrir_RecordsetSeaMetal(RDatos, "Select Head, StandardValueId, Avg(MeasurementValue) From ReportValues Where ReportId = " & TxtDoc.Text & " Group By Head, StandardValueId")
        
    If RDatos.RecordCount > 0 Then
    Else
        MsgBox "Documento No Existe En EPA", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
    End If
    
    
            Do Until RDatos.EOF
                '0 ESPESOR - RUTINA 0
                If RDatos!StandardValueId = "0" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Espesor " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '1 DIAMETRO EXTERIOR - RUTINA 1
                ElseIf RDatos!StandardValueId = "1" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Diametro Exterior " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '2 DIAMETRO INTERIOR - RUTINA 2
                ElseIf RDatos!StandardValueId = "2" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '2' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Diametro Interior " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '3 DIAMETRO CHUCK - RUTINA 3
                ElseIf RDatos!StandardValueId = "3" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Diametro Chuck " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '4 VUELO - RUTINA 4
                ElseIf RDatos!StandardValueId = "4" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '4' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Vuelo " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '5 ALTURA REBORDEADOR - RUTINA 5
                ElseIf RDatos!StandardValueId = "5" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '5' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Altura Rebordeador " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '6 RADIO EXTERIOR - RUTINA 6
                ElseIf RDatos!StandardValueId = "6" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '6' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Radio Exterior " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '7 PROFUNDIDAD - RUTINA 7
                ElseIf RDatos!StandardValueId = "7" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Profundidad " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '-8 ANCHO DE CANAL - RUTINA -8
                ElseIf RDatos!StandardValueId = "-8" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '-8' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Ancho De Canal " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                '8 PRIMER RELIEVE - RUTINA 8
                ElseIf RDatos!StandardValueId = "8" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '8' and Cabezal = " & RDatos!Head
                                    If Err <> 0 Then
                                        MsgBox "Error En Medida Primer Relieve " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                End If
                
                RDatos.MoveNext
            Loop
    
    MousePointer = 0
        
        
    MsgBox "Rutinas Capturadas Con Exito", vbOKOnly + vbInformation, "Informacion"
    
    
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DbGridSeametal_HeadClick(ByVal ColIndex As Integer)
        RSeaMetal.Sort = RSeaMetal.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
On Error Resume Next

        If GPlanta = "CULIACAN" Then
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaEpa
        ElseIf GPlanta = "SAN LUIS" Then
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaEpaSanLuisPotosi
        ElseIf GPlanta = "CHIAPAS" Then
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaEpaChiapas
        End If
            
    
            'CONEXION A SEAMETAL
            Set ConexionSeametal = New ADODB.Connection
            ConexionSeametal.ConnectionString = GConectionStringSeaMetal
            ConexionSeametal.Open
            
            If Err <> 0 Then
                MsgBox "Error Al Hacer La Conexion con BaseDeDatos " & Err.Number & " " & Err.Description
                Exit Sub
            End If
                
   
    
    'SELECCIONA LOS CAMPOS DE LA BASE DE DATOS
    Set RSeaMetal = New ADODB.Recordset
    Call Abrir_RecordsetSeaMetal(RSeaMetal, "Select R.ReportId, R.Head, R.standardValueID, N.Nombre, avg(R.MeasurementValue/1000) From ReportValues R, NombresMedidas N Where R.standardValueID = N.standardValueIndex group by R.ReportId, R.Head, R.standardValueID, N.Nombre")
    Set DbGridSeametal.DataSource = RSeaMetal
    
    If Err <> 0 Then
            MsgBox "Error Al Seleccionar Datos De La Tabla ReportValues " & Err.Number & " " & Err.Description, vbInformation, "Error"
            'Exit Sub
    End If
    
    DbGridSeametal.Columns(0).Caption = "Reporte"
    DbGridSeametal.Columns(1).Caption = "Cabezal"
    DbGridSeametal.Columns(2).Caption = "Codigo"
    DbGridSeametal.Columns(3).Caption = "Medida"
    DbGridSeametal.Columns(4).Caption = "Valor"
    DbGridSeametal.Columns(2).Width = "500"
    DbGridSeametal.Columns(4).Width = "1000"
    DbGridSeametal.Columns(2).Alignment = dbgRight
    DbGridSeametal.Columns(4).Alignment = dbgRight
    DbGridSeametal.Columns(4).NumberFormat = "#,###,##0.00"
    RSeaMetal.MoveLast
    MskFec.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ConexionSeametal.Close
    Set ConexionSeametal = Nothing
    If Err <> 0 Then
    End If
End Sub

Private Sub MskFec_GotFocus()
        MskFec.SelStart = 0
        MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub MskFec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab"
        End If
End Sub

Private Sub TxtDoc_GotFocus()
        TxtDoc.SelStart = 0
        TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab"
        End If
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab"
        End If
End Sub

Private Sub TxtHor_GotFocus()
        TxtHor.SelStart = 0
        TxtHor.SelLength = Len(TxtHor.Text)
End Sub

Private Sub TxtHor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab"
        End If
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab"
        End If
End Sub

Private Sub TxtLin_LostFocus()
On Error Resume Next
    Set RDatosdeRutina = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RDatosdeRutina, "Select Fec_Rut, Esp_Tec, Hor_Rut From CapturaRutinas Where Linea = '" & TxtLin.Text & "' And Fec_Rut >= #" & Format((Date - 1), "mm/dd/yyyy") & "# Order By Fec_rut, Hor_Rut")
        Else
            Call Abrir_Recordset(RDatosdeRutina, "Select Fec_Rut, Esp_Tec, Hor_Rut From CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And Fec_Rut >= To_Date('" & (Date - 1) & "', 'dd/mm/yyyy') Order By Fec_rut, Hor_Rut")
        End If
    If Err <> 0 Then
        MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    If RDatosdeRutina.RecordCount > 0 Then
        RDatosdeRutina.MoveLast
        MskFec.Text = RDatosdeRutina(0)
        TxtFicTec.Text = RDatosdeRutina(1)
        TxtHor.Text = RDatosdeRutina(2)
    End If
    
End Sub
