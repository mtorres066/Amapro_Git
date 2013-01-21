VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraCapturaRutinas 
   Caption         =   "Genera Captura de Rutinas AUTOMATICA"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   Icon            =   "GeneraCapturaRutinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DbGridSeametal 
      Height          =   5895
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10398
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
      Left            =   4800
      TabIndex        =   16
      Top             =   6600
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
      Top             =   6720
      Width           =   855
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   6720
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
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox TxtFicTec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   6720
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
      Left            =   9120
      Picture         =   "GeneraCapturaRutinas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1695
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
      Left            =   7080
      Picture         =   "GeneraCapturaRutinas.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"GeneraCapturaRutinas.frx":4526
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   10860
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "Doc."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Hora"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Ficha Tecnica"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Linea"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   6720
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   6600
      Width           =   10935
   End
End
Attribute VB_Name = "GeneraCapturaRutinas"
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
    Call Abrir_RecordsetSeaMetal(RDatos, "Select HeadNo, Avg(Thick), Avg(Length), Avg(Counter), avg(CoverH), avg(BodyH), Avg(Overlap), Avg(Ovlp), Avg(SeamGap), sum(Tightness) From Report_Values Where ReportId = " & TxtDoc.Text & " Group By HeadNo")
        
    If RDatos.RecordCount > 0 Then
    Else
        MsgBox "Documento No Existe En SEAMETAL", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
    End If
    
    
            Do Until RDatos.EOF
        
'TIPO DE ENVASE
        
'2 PIEZAS ____________________________________________________________________________________________________
        
                If OptDos.Value = True Then
                     If Not IsNull(RDatos(7)) Then
                        '1002 ALTURA TERMINADO 2 PIEZAS
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1002' and Cabezal = " & RDatos!HeadNo
                            End If
                                If Err <> 0 Then
                                    MsgBox Err.Description
                                    Err.Clear
                                End If
                    End If
            
'3 PIEZAS ____________________________________________________________________________________________________
        
                ElseIf Opttres.Value = True Then
                
                    If Not IsNull(RDatos(0)) Then
                        '2000 ANCHO DE CIERRE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '2000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '2000' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(1)) Then
                        '4000 LARGO DE CIERRE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '4000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '4000' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(2)) Then
                        '3000 PROFUNDIDAD
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '3000' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(3)) Then
                        '6000 GANCHO DE FONDO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '6000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '6000' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(4)) Then
                        '5000 GANCHO DE CUERPO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '5000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '5000' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(5)) Then
                        '7000 TRASLAPE
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7000' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7000' and Cabezal = " & RDatos!HeadNo
                        End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(6)) Then
                        If RDatos(6) > 0 Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                '7250 % DE TRASLAPE  (solo en mexico, en guatemala no se mide automatico
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7250' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7250' and Cabezal = " & RDatos!HeadNo
                            End If
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                    Err.Clear
                                End If
                        End If
                    End If
                    
                    If Not IsNull(RDatos(7)) Then
                        If RDatos(7) > 0 Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                '1000 ALTURA TERMINADO (solo en mexico, en guatemala no se mide automatico desde el seametal)
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1000' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1000' and Cabezal = " & RDatos!HeadNo
                            End If
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                    Err.Clear
                                End If
                        End If
                    End If
                    
        
                    If Not IsNull(RDatos(8)) Then
                        '7500 % APLANCHADO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(9), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7500' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(9) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7500' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
        
'ANILLO ____________________________________________________________________________________________________
        
                ElseIf OptAnillo.Value = True Then
                
                    If Not IsNull(RDatos(0)) Then
                        '2001 ANCHO DE CIERRE
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '2001' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '2001' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(1)) Then
                        '4001 LARGO DE CIERRE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '4001' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '4001' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(2)) Then
                        '3001 PROFUNDIDAD
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3001' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '3001' and Cabezal = " & RDatos!HeadNo
                            End If
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Err.Clear
                        End If
                    End If
                        
                    If Not IsNull(RDatos(3)) Then
                        '6001 GANCHO DE FONDO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '6001' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '6001' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(4)) Then
                        '5001 GANCHO DE CUERPO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '5001' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '5001' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(5)) Then
                        '7001 TRASLAPE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7001' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7001' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(6)) Then
                        '7250 % DE TRASLAPE UNICO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7250' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7250' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                    
                    If Not IsNull(RDatos(7)) Then
                        '1002 ALTURA TERMINADO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1002' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                    
        
                    If Not IsNull(RDatos(8)) Then
                        '7501 % APLANCHADO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(9), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7501' and Cabezal = " & RDatos!HeadNo
                            Else
                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(9), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7501' and Cabezal = " & RDatos!HeadNo
                            End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                
'FONDO ____________________________________________________________________________________________________
        
                ElseIf OptFondo.Value = True Then
                
                    If Not IsNull(RDatos(0)) Then
                        '2002 ANCHO DE CIERRE
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '2002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(1) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '2002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(1)) Then
                        '4002 LARGO DE CIERRE
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '4002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(2) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '4002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(2)) Then
                        '3002 PROFUNDIDAD
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(3) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '3002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(3)) Then
                        '6002 GANCHO DE FONDO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '6002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(4) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '6002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(4)) Then
                        '5002 GANCHO DE CUERPO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '5002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(5) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '5002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(5)) Then
                        '7002 TRASLAPE
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(6) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                        
                    If Not IsNull(RDatos(6)) Then
                        '7250 % DE TRASLAPE UNICO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7250' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(7), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7250' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                    
                    If Not IsNull(RDatos(7)) Then
                        '1002 ALTURA TERMINADO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(8) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1002' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
        
                    If Not IsNull(RDatos(8)) Then
                        '7502 % APLANCHADO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(9), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '7502' and Cabezal = " & RDatos!HeadNo
                        Else
                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format(RDatos(9), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '7502' and Cabezal = " & RDatos!HeadNo
                        End If
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
                
                
                End If 'FIN DE OPCION DE 2 O 3 PIEZAS ANILLO Y FONDO
                    
                RDatos.MoveNext
            Loop
    
    MousePointer = 0
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
        
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
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaSeametal
        ElseIf GPlanta = "SAN LUIS" Then
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaSeametalSanLuisPotosi
        ElseIf GPlanta = "CHIAPAS" Then
            GConectionStringSeaMetal = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaSeametalChiapas
        End If
            
    
            'CONEXION A SEAMETAL
            Set ConexionSeametal = New ADODB.Connection
            ConexionSeametal.ConnectionString = GConectionStringSeaMetal
            ConexionSeametal.Open
            
            If Err <> 0 Then
                MsgBox "Error Al Hacer La Conexion Con Base De Datos Seametal " & Err.Number & " " & Err.Description
                Exit Sub
            End If
                
    'ASIGNA LA UBICACION EN LA RED DE LA MAQUINA DE SEAMETAL
    'Set DBSeaMetal = OpenDatabase("\\Seametal\SeaMetal\SeaMetal.MDB", False, False)
    'Set DBSeaMetal = OpenDatabase("C:\SEAMETAL\seametal.mdb", False, False)
    
        'If Err <> 0 Then
        'End If
        
    'ASIGNA LA UBICACION EN LA RED DE LA MAQUINA DE SEAMETAL
    'DataSeaMetal.DatabaseName = "\\Seametal\Seametal\Seametal.mdb"
    'DataSeaMetal.DatabaseName = "c:\Seametal\Seametal.mdb"
    
    
    'SELECCIONA LOS CAMPOS DE LA BASE DE DATOS
    Set RSeaMetal = New ADODB.Recordset
    Call Abrir_RecordsetSeaMetal(RSeaMetal, "Select ReportId, HeadNo, CheckDate, Thick, Length, Counter, CoverH, BodyH, Overlap, Ovlp, SeamGap, Tightness From Report_Values Order By ReportId, HeadNo")
    Set DbGridSeametal.DataSource = RSeaMetal
    
    If Err <> 0 Then
            MsgBox "Error Al Selecciona Datos De La Tabla ReportValues " & Err.Number & " " & Err.Description, vbInformation, "Error"
            'Exit Sub
    End If
    
    DbGridSeametal.Columns(2).Width = "2000"
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
