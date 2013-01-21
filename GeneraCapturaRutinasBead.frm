VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraCapturaRutinasCpa 
   Caption         =   "Genera Captura de Rutinas Automatica CPA"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   Icon            =   "GeneraCapturaRutinasBead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DbGridCpa2 
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
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
   Begin MSDataGridLib.DataGrid DbGridCpa1 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
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
      Left            =   4680
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
      Begin VB.OptionButton Opttres 
         Caption         =   "3 Piezas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptDos 
         Caption         =   "2 Piezas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   7080
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
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox TxtFicTec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox TxtLin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   6960
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
      Height          =   855
      Left            =   9240
      Picture         =   "GeneraCapturaRutinasBead.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1575
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
      Height          =   855
      Left            =   7080
      Picture         =   "GeneraCapturaRutinasBead.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label LblCpa2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0500  Ancho Pestaña Sup.   0600  Ancho Pestaña Inf.  1002 Altura    3001  Profundidad Counter    3002  Profundidad Centro"
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
      TabIndex        =   16
      Top             =   6480
      Width           =   10815
   End
   Begin VB.Label LblCpa1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0800 Profundidad De Acanalado  0900 Ancho De Acanalado   0501 Ancho Pestaña 2 Piezas  1001 Altura Entre Pest. C/Acan"
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
      TabIndex        =   15
      Top             =   6240
      Width           =   10815
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Hora"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Ficha Tecnica"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Linea"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   6840
      Width           =   10935
   End
End
Attribute VB_Name = "GeneraCapturaRutinasCpa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaRutina As New ADODB.Recordset
Dim RBuscaCabezalesRutina As New ADODB.Recordset
Dim RBuscaDatoas As New ADODB.Recordset
Dim RMaximo As New ADODB.Recordset
Dim RDatos As New ADODB.Recordset
Dim RCpa1 As New ADODB.Recordset
Dim RCpa2 As New ADODB.Recordset
Dim Cont As Integer



Private Sub CmdCapturar_Click()
On Error Resume Next
MousePointer = 11

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
    
    'BUSCA EL MAXIMO
    Set RMaximo = New ADODB.Recordset
    Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultBeads")
        
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        'DOS PIEZAS __________________________________________________________________________________________________
        
        If OptDos.Value = True Then
                            '0501 ANCHO DE PESTAÑA INFERIOR_____________________________________________________________________
                            
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '0501'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '0501'")
                                End If
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                    Set RDatos = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RDatos, "Select avg(BottomFlangeMeasurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '0501 ANCHO DE PESTAÑA SUPERIOR
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0501' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '0501' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0501' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
                            
                            
                            '1002 ALTURA TERMINADA ____________________________________________________________________
                            
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '1002'")
                                Else
                                        Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '1002'")
                                End If
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                        Set RDatos = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RDatos, "Select avg(CanHeightMeasurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '1002 ANCHO DE PESTAÑA SUPERIOR
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1002' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                                
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
                            
                            
                            '3001 PROFUNIDAD COUNTER ________________________________________________________________
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '3001'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '3001'")
                                End If
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                    Set RDatos = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RDatos, "Select avg(ExtDevice1Measurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '3001 PROFUNDIDAD COUNTER
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3001' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '3001' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3001' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
                            
                            
                            
                            '3002 PROFUNIDAD CENTRO ________________________________________________________________
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '3002'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '3002'")
                                End If
                                
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                    Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                    Set RDatos = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RDatos, "Select avg(ExtDevice2Measurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '3002 PROFUNDIDAD COUNTER
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3002' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '3002' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '3002' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
                            
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        'TRES PIEZAS __________________________________________________________________________________________________
        
        Else
                        '0800 PROFUNDIDAD DE ACANALADO _____________________________________________________________________
                            
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '0800'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '0800'")
                                End If
                                If RBuscaCabezalesRutina.RecordCount > 0 Then
                                        Cont = 1
                                        Do Until Cont > RBuscaCabezalesRutina!cabezal
                                        'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                        'saca el promedio de las medias en base al cabezal
                                        Set RDatos = New ADODB.Recordset
                                            Call Abrir_RecordsetCpa(RDatos, "Select avg(Depth) From MeasurementResultBeads Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                                If RDatos.RecordCount > 0 Then
                                                    If Not IsNull(RDatos(0)) Then
                                                            '0800 PROFUNDIDAD DE ACANALADO
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0800' and Cabezal = " & Cont
                                                            Else
                                                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = TO_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '0800' and Cabezal = " & Cont
                                                            End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                            'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1002' and Cabezal = " & Cont)
                                                            'If RBuscaRutina.RecordCount > 0 Then
                                                            '        RBuscaRutina.Edit
                                                            '                RBuscaRutina!Valor = Format((RDatos(1) / 1000), "#,###,##0.00")
                                                            '        RBuscaRutina.Update
                                                            'End If
                                                    End If
                                                End If
                                                    Cont = Cont + 1
                                        Loop
                                Else
                                End If
                                        
                        '0900 ANCHO DE ACANALADO____________________________________________________________________________
                            
                            'Set RBuscaCabezalesRutina = New ADODB.Recordset
                            '    If GOrigenDeDatos = "AmaproAccess" Then
                            '        Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '0900'")
                            '    Else
                            '        Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '0900'")
                            '    End If
                            'If RBuscaCabezalesRutina.RecordCount > 0 Then
                            '        Cont = 1
                            '        Do Until Cont > RBuscaCabezalesRutina!Cabezal
                            '        'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                            '        'saca el promedio de las medias en base al cabezal
                            '        Set RDatos = New ADODB.Recordset
                            '            'ESTA RUTINA EXISTIA ANTES, PERO EN YA NO ESTA ESTE CAMPO, COMO QUE EN LA VERSION ULTIMA
                                        'LO QUITO QUALITY
                            '            Call Abrir_RecordsetCpa(RDatos, "Select avg(Width) From MeasurementResultBeads Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                            '                If RDatos.RecordCount > 0 Then
                            '                    If Not IsNull(RDatos(0)) Then
                            '                            '0900 ANCHO DE ACANALADO
                            '                            If GOrigenDeDatos = "AmaproAccess" Then
                            '                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0900' and Cabezal = " & Cont
                            '                            Else
                            '                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '0900' and Cabezal = " & Cont
                            '                            End If
                            '                                    If Err <> 0 Then
                            '                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            '                                        Err.Clear
                            '                                    End If
                            '                    End If
                            '                End If
                            '                    Cont = Cont + 1
                            '        Loop
                            'Else
                            'End If
                                
                            
                        '0600 ANCHO DE PESTAÑA INFERIOR___________________________________________________________________
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '0600'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '0600'")
                                End If
                                
                                                        
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                                        'BUSCA EL MAXIMO
                                        Set RMaximo = New ADODB.Recordset
                                            Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                        Cont = 1
                                        
                                        Do Until Cont > RBuscaCabezalesRutina!cabezal
                                        'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                        'saca el promedio de las medias en base al cabezal
                                        Set RDatos = New ADODB.Recordset
                                            Call Abrir_RecordsetCpa(RDatos, "Select avg(BottomFlangeMeasurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                                If Not IsNull(RDatos(0)) Then
                                                            '0600 ANCHO DE PESTAÑA INFERIOR
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0600' and Cabezal = " & Cont
                                                            Else
                                                                Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '0600' and Cabezal = " & Cont
                                                            End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                            'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0600' and Cabezal = " & Cont)
                                                            'If RBuscaRutina.RecordCount > 0 Then
                                                            '    RBuscaRutina.Edit
                                                            '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                            '    RBuscaRutina.Update
                                                            'End If
                                                End If
                                                    
                                        
                                            Cont = Cont + 1
                                        Loop
                            End If
                            
                        '0500 ANCHO DE PESTAÑA SUPERIOR_____________________________________________________________________
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '0500'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '0500'")
                                End If
                                                        
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                            
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                        Set RDatos = New ADODB.Recordset
                                            Call Abrir_RecordsetCpa(RDatos, "Select avg(TopFlangeMeasurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '0500 ANCHO DE PESTAÑA SUPERIOR
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0500' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '0500' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '0500' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
                        '1001 ALTURA ENTRE PESTAÑAS CON ACANALADO __________________________________________________________________________
                            Set RBuscaCabezalesRutina = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where Rutina = '1001'")
                                Else
                                    Call Abrir_Recordset(RBuscaCabezalesRutina, "Select Cabezal From Rutinas Where UPPER(Rutina) = '1001'")
                                End If
                                                           
                            If RBuscaCabezalesRutina.RecordCount > 0 Then
                            
                                    'BUSCA EL MAXIMO
                                    Set RMaximo = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RMaximo, "Select Max(MeasurementId) From MeasurementResultItems")
                                    Cont = 1
                                    
                                    Do Until Cont > RBuscaCabezalesRutina!cabezal
                                    'SELECCIONAMOS LOS DATOS QUE ESTAN CON ESTE REPORTE
                                    'saca el promedio de las medias en base al cabezal
                                    Set RDatos = New ADODB.Recordset
                                        Call Abrir_RecordsetCpa(RDatos, "Select avg(CanHeightMeasurement) From MeasurementResultItems Where MeasurementId = " & RMaximo(0) & " and Head = " & Cont)
                                            If Not IsNull(RDatos(0)) Then
                                                        '1001 ALTURA ENTRE PESTAÑAS CON ACANALADO
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where Linea = '" & TxtLin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1001' and Cabezal = " & Cont
                                                        Else
                                                            Conexion.Execute "Update CapturaRutinas Set Valor = " & Format((RDatos(0) / 1000), "#,###,##0.00") & " Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and UPPER(Esp_tec) = '" & UCase(TxtFicTec.Text) & "' and Fec_Rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_Rut) = '" & UCase(TxtHor.Text) & "' and UPPER(Rutina) = '1001' and Cabezal = " & Cont
                                                        End If
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                    Err.Clear
                                                                End If
                                                        
                                                        'Set RBuscaRutina = Db.OpenRecordset("Select Valor from CapturaRutinas Where Linea = '" & Txtlin.Text & "' and Esp_tec = '" & TxtFicTec.Text & "' and Fec_Rut = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_Rut = '" & TxtHor.Text & "' and Rutina = '1001' and Cabezal = " & Cont)
                                                        'If RBuscaRutina.RecordCount > 0 Then
                                                        '    RBuscaRutina.Edit
                                                        '            RBuscaRutina!Valor = Format((RDatos(0) / 1000), "#,###,##0.00")
                                                        '    RBuscaRutina.Update
                                                        'End If
                                            End If
                                        Cont = Cont + 1
                                    Loop
                            End If
        End If 'FIN DE TIPO DE ENVASE
    
'___________________________________________________________________________________________________
    
                    
    MousePointer = 0
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
    End If
    
    MsgBox "Captura de Rutinas Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
    
    
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DbGridCpa1_HeadClick(ByVal ColIndex As Integer)
        RCpa1.Sort = RCpa1.Fields(ColIndex).Name
End Sub

Private Sub DbGridCpa2_HeadClick(ByVal ColIndex As Integer)
        RCpa2.Sort = RCpa2.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
On Error Resume Next
        If GPlanta = "CULIACAN" Then
            GConectionStringCpa = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaCpa
        ElseIf GPlanta = "SAN LUIS" Then
            GConectionStringCpa = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaCpaSanLuisPotosi
        ElseIf GPlanta = "CHIAPAS" Then
            GConectionStringCpa = "Provider=Microsoft.Jet.OLEDB.3.51;User ID=Admin;Data Source=" & GRutaCpaChiapas
        End If
        
            

            'CONEXION A CPA
            Set ConexionCpa = New ADODB.Connection
            ConexionCpa.ConnectionString = GConectionStringCpa
            ConexionCpa.Open
            
    'ASIGNA LA UBICACION EN LA RED DE LA MAQUINA DE SEAMETAL
    'Set DBCpa = OpenDatabase("\\Seametal\Cpa\Bead.MDB", False, False)
    'Set DBCpa = OpenDatabase("c:\Cpa\Bead.MDB", False, False)
    
        
            If Err <> 0 Then
               MsgBox "Error Al Hacer La Conexion Con Base De Datos CPA " & Err.Number & " " & Err.Description
               Exit Sub
            End If
            
        
        
    'ASIGNA LA UBICACION EN LA RED DE LA MAQUINA DE SEAMETAL
    'DataCPA.DatabaseName = "\\Seametal\Cpa\Bead.mdb"
    'DataCPA.DatabaseName = "c:\Cpa\Bead.mdb"
    'DataCpa2.DatabaseName = "\\Seametal\Cpa\Bead.mdb"
    'DataCpa2.DatabaseName = "c:\Cpa\Bead.mdb"
    
        
    'SELECCIONA LOS CAMPOS DE LA BASE DE DATOS
    Set RCpa1 = New ADODB.Recordset
    Call Abrir_RecordsetCpa(RCpa1, "Select * From MeasurementResultBeads")
    Set DbGridCpa1.DataSource = RCpa1
    
    If Err <> 0 Then
            MsgBox "Error Al Seleccionar Datos ReportValues " & Err.Number & " " & Err.Description, vbInformation, "Error"
            
    End If
    
    DbGridCpa1.Columns(2).Width = "2000"
    RCpa1.MoveLast
    
    
    
    Set RCpa2 = New ADODB.Recordset
    Call Abrir_RecordsetCpa(RCpa2, "Select * From MeasurementResultItems")
    Set DbGridCpa2.DataSource = RCpa2
    
    If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbInformation, "Error"
    End If
    
    DbGridCpa2.Columns(2).Width = "2000"
    RCpa2.MoveLast
    
    
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ConexionCpa.Close
    Set ConexionCpa = Nothing
    If Err <> 0 Then
        
    End If
End Sub

Private Sub MskFec_GotFocus()
        MskFec.SelStart = 0
        MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
        
End Sub

Private Sub TxtHor_GotFocus()
        TxtHor.SelStart = 0
        TxtHor.SelLength = Len(TxtHor.Text)
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_LostFocus()
    Set RDatos = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RDatos, "Select Fec_Rut, Esp_Tec, Hor_Rut From CapturaRutinas Where Linea = '" & TxtLin.Text & "' And Fec_Rut >= #" & Format((Date - 1), "mm/dd/yyyy") & "# Order By Fec_rut, Hor_Rut")
        Else
            'Call Abrir_Recordset(RDatos, "Select Fec_Rut, Esp_Tec, Hor_Rut From CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And To_Char(Fec_Rut, 'yyyy') = " & Year(Date) & " And To_Char(Fec_Rut, 'mm') = " & Month(Date) & " Order By Fec_rut, Hor_Rut")
            Call Abrir_Recordset(RDatos, "Select Fec_Rut, Esp_Tec, Hor_Rut From CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And Fec_Rut >= To_Date('" & (Date - 1) & "', 'dd/mm/yyyy') Order By Fec_rut, Hor_Rut")
        End If
    If RDatos.RecordCount > 0 Then
        RDatos.MoveLast
    End If
    
    If RDatos.RecordCount > 0 Then
        MskFec.Text = RDatos(0)
        TxtFicTec.Text = RDatos(1)
        TxtHor.Text = RDatos(2)
    End If
    
    
End Sub

