VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ActividadesDiariasConsulta 
   BackColor       =   &H0080C0FF&
   Caption         =   "Consulta Actividades Diarias"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ActividadesDiariasConsulta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   8295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   7215
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "ActividadesDiariasConsulta.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   7095
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
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
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Height          =   855
      Left            =   9480
      Picture         =   "ActividadesDiariasConsulta.frx":3D6C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.TextBox TxtUbi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   82903043
      CurrentDate     =   38483
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   82903043
      CurrentDate     =   38483
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   855
      Left            =   10800
      Picture         =   "ActividadesDiariasConsulta.frx":41F4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdVerDatos 
      Caption         =   "&Consultar"
      Height          =   855
      Left            =   8400
      Picture         =   "ActividadesDiariasConsulta.frx":470F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ver Datos"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LblUbi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Empleado"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblUbiDes 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Hasta"
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Desde"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "ActividadesDiariasConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RVerDatos As New ADODB.Recordset
Dim RDetalleOrden As New ADODB.Recordset

Dim RBusqueda As New ADODB.Recordset

Dim RBuscaEmpleado As New ADODB.Recordset
Dim RBuscaDepartamento As New ADODB.Recordset
Dim RBuscaSeccion As New ADODB.Recordset
Dim RBuscaMaquina As New ADODB.Recordset
Dim RBuscaSistema As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset
Dim RBuscaUltimaOrden As New ADODB.Recordset
Dim RBuscaTareas As New ADODB.Recordset

Dim RBuscaSolicitud As New ADODB.Recordset
Dim RBuscaEmpleado2 As New ADODB.Recordset

Dim RBuscaOrden As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset

Dim BEmpresa As Boolean
Dim BDepartamento As Boolean
Dim BSeccion As Boolean
Dim BMaquina As Boolean
Dim BSistema As Boolean
Dim BEquipo As Boolean

Dim VDocumento As Long
Dim vtexto As String
Dim VContador As Integer
Dim VRespuesta As String



Private Sub CmdImprimir_Click()
On Error Resume Next
                    GCriteriaReporte = "{M_ActividadesDiarias.Fecha} >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And {M_ActividadesDiarias.Fecha} <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And {M_ActividadesDiarias.Empleado} Like '" & UCase(TxtUbi.Text) & "*'"
                    GTituloReporte = "Desde " & DTPFecIni.Value & " Hasta " & DtpFecFin.Value & " Y Empleado "
                    GNombreReporte = "DeMante_ActividadesDiariasO.rpt"
                    FrmReporte.Show 1
                                                
                If Err <> 0 Then
                End If
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub CmdVerDatos_Click()
    verdatos
End Sub



Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
                RVerDatos.Sort = RVerDatos.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

End Sub

Private Sub DbGridBusqueda_DblClick()
            TxtUbi.Text = DBGridBusqueda.Columns(0).Text
            TxtUbi.SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DbGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                            TxtUbi.Text = DBGridBusqueda.Columns(0).Text
                            TxtUbi.SetFocus
                            FrameBusqueda.Visible = False
            End If

End Sub

Private Sub Form_Load()
        DTPFecIni.Value = Date
        DtpFecFin.Value = Date
        verdatos
End Sub





Private Sub Form_Resize()
On Error Resume Next
    DataGrid1.Width = (Me.Width - 300)
    DataGrid1.Height = (Me.Height - 1700)
    DataGrid1.Columns(5).Width = (Me.Width - 6700)
    
    If Err <> 0 Then
    End If
End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme where Nombre Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme where UPPER(Nombre) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme where No_Emple Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme where UPPER(No_Emple) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            
                    
                    If RBusqueda.RecordCount > 0 Then
                    End If
                    
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub TxtUbi_Change()
            'EMPLEADO
            
                    Set RBuscaEmpleado = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where No_Emple = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where UPPER(No_Emple) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaEmpleado.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaEmpleado!Nombre
                        Else
                            LblUbiDes.Caption = ""
                        End If
            

End Sub

Private Sub TxtUbi_DblClick()
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus
        
        

End Sub

Private Sub TxtUbi_GotFocus()
        TxtUbi.SelStart = 0
        TxtUbi.SelLength = Len(TxtUbi.Text)
End Sub

Private Sub TxtUbi_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                                Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From Arplme")
                                Set DBGridBusqueda.DataSource = RBusqueda
                                DBGridBusqueda.Columns(1).Width = "4000"
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus
                    End If

End Sub



Public Sub verdatos()
On Error Resume Next
MousePointer = 11
        Set RVerDatos = New ADODB.Recordset
                
                            vtexto = "Select A.Documento, A.Fecha, E.Nombre, M.Descripcion, A.Minutos, A.Descripcion as Actividad From M_ActividadesDiarias A, M_TiposMantenimiento M, Arplme E Where A.Fecha >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " And A.Fecha <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy') And A.Empleado Like '" & TxtUbi.Text & "%' And A.TipoMantenimiento = M.Codigo And A.Empleado = E.No_emple and A.EmpresaEmpleado = E.No_Cia"
                        
                        
                    Call Abrir_Recordset(RVerDatos, vtexto)
                    Set DataGrid1.DataSource = RVerDatos
                    
                            
                            DataGrid1.Columns(0).Width = "10"
                            DataGrid1.Columns(1).Width = "1000"
                            DataGrid1.Columns(2).Width = "2000"
                            DataGrid1.Columns(3).Width = "2000"
                            DataGrid1.Columns(4).Width = "500"
                            DataGrid1.Columns(5).Width = "5300"
                            
                            DataGrid1.Columns(0).Caption = "Documento"
                            DataGrid1.Columns(1).Caption = "Fecha"
                            DataGrid1.Columns(2).Caption = "Empleado"
                            DataGrid1.Columns(3).Caption = "Mantenimiento"
                            DataGrid1.Columns(4).Caption = "Minutos"
                            DataGrid1.Columns(5).Caption = "Descripcion Del Trabajo"
                            
                            
                            'DataGrid1.Columns(5).AllowSizing = True
                            DataGrid1.Columns(2).WrapText = True
                            DataGrid1.Columns(5).WrapText = True
                            DataGrid1.RowHeight = "800"
                            
                            DataGrid1.Columns(1).NumberFormat = "dd/mm/yyyy"
                            'DataGrid1.Columns(4).Alignment = dbgRight

                    
            MousePointer = 0
            
                If Err <> 0 Then
                    MsgBox "ERROR " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If

            

End Sub
