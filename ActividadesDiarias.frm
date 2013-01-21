VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ActividadesDiarias 
   Caption         =   "Actividades Diarias"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "ActividadesDiarias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4935
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8705
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
      Begin VB.CommandButton CmdSaleBusqueda 
         Caption         =   "Salida"
         Height          =   735
         Left            =   6480
         Picture         =   "ActividadesDiarias.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "salida"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
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
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
   End
   Begin VB.Frame FrameEncabezado 
      Height          =   6135
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Txttexto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2565
         Index           =   7
         Left            =   1560
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "250 caracteres maximo"
         Top             =   2400
         Width           =   6105
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   975
         Left            =   1560
         Picture         =   "ActividadesDiarias.frx":3D6C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox Txttexto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   2
         Top             =   960
         Width           =   1260
      End
      Begin VB.TextBox Txttexto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2040
         Width           =   1260
      End
      Begin VB.TextBox Txttexto 
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
         Index           =   6
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Doble click o signo '+' para ayuda"
         Top             =   600
         Width           =   1260
      End
      Begin MSMask.MaskEdBox MskFec 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskParFin 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskParIni 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Final"
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
         TabIndex        =   24
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
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
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion Del Trabajo"
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empleado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mantenimiento"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Minutos"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label LblMan 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label LblEmpl 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   600
         Width           =   4815
      End
   End
End
Attribute VB_Name = "ActividadesDiarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaTipoTrabajo As New ADODB.Recordset
Dim RInventario As New ADODB.Recordset

Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaClienteNit As New ADODB.Recordset
Dim RBuscaTransaccion As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset

Dim RBuscaInvExi As New ADODB.Recordset
Dim RBuscaDocumentoMaximo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBusqueda2 As New ADODB.Recordset
Dim RBuscaTipo As New ADODB.Recordset

Dim RBuscaEmpleado As New ADODB.Recordset
Dim RBuscaEmpresa As New ADODB.Recordset
Dim RBuscaMantenimiento As New ADODB.Recordset
Dim RBuscaSeccion As New ADODB.Recordset
Dim RBuscaMaquina As New ADODB.Recordset
Dim RBuscaSistema As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset

Dim BEmpresa As Boolean
Dim BMantenimiento As Boolean
Dim BSeccion As Boolean
Dim BMaquina As Boolean
Dim BSistema As Boolean
Dim BEquipo As Boolean
Dim BEmpleado As Boolean
Dim BTipoTrabajo As Boolean

Dim Cont As Integer
Dim mensaje As String
Dim VPrecioTotal As Currency
Dim VDocumento As String

Dim VTotal As Currency
'Dim VLetras As String
Dim VTipo As String

Dim VCliente As String
Dim VDocumentoMaximo As Long
Dim VFecha As Date

Dim VCantidad As Currency
Dim VPrecioUnitario As Currency
Dim VCodigo As String
Dim vtexto As String
Dim VMonto As Currency
Dim VDescuento As Currency
Dim VDescuento2 As Currency

Dim VHoraInicial As Date
Dim VHoraFinal As Date

Private Sub CmdGrabar_Click()
On Error Resume Next

MousePointer = 11

    If Not IsDate(MskFec.Text) Then
        MousePointer = 0
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Txttexto.Item(2).Text) Then
        MousePointer = 0
        MsgBox "Minutos Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        Txttexto.Item(2).SetFocus
        Exit Sub
    End If
    
    
    'VERIFICA LAS OBSERVACIONES
    If Txttexto.Item(7).Text = "" Then
        MsgBox "El Campo De Descripcion Del Trabajo No Puede Estar Vacio, Escriba Alguna Observacion ", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    Txttexto.Item(1).Text = UCase(Txttexto.Item(1).Text)
    
        
                
                Set RBuscaMantenimiento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaMantenimiento, "Select Descripcion From M_TiposMantenimiento Where Codigo = '" & Txttexto.Item(1).Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaMantenimiento, "Select Descripcion From M_TiposMantenimiento Where UPPER(Codigo) = '" & UCase(Txttexto.Item(1).Text) & "'")
                    End If
                    If RBuscaMantenimiento.RecordCount > 0 Then
                        
                    Else
                        MsgBox "Departamento Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(1).SetFocus
                        Exit Sub
                    End If
        
                
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where No_Emple = '" & Txttexto.Item(6).Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where UPPER(No_Emple) = '" & UCase(Txttexto.Item(6).Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        
                    Else
                        MsgBox "Empleado Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(6).SetFocus
                        Exit Sub
                    End If
            
            
    MousePointer = 0
    
    mensaje = MsgBox("Desea Grabar?", vbYesNo + vbInformation + vbDefaultButton2, "Confirmacion")
        
    If mensaje = vbYes Then
    Else
        Exit Sub
    End If
    
    MousePointer = 11
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    Set RBuscaSigDoc = New ADODB.Recordset
    Call Abrir_Recordset(RBuscaSigDoc, "Select Max(Documento) from M_ActividadesDiarias")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                VDocumento = "1"
            Else
                VDocumento = RBuscaSigDoc(0) + 1
            End If
        End If
        
    
        'INICIA LA TRANSACCION
        Conexion.BeginTrans
    
                    'ENCABEZADO DE TRANSACCIONES
                    If GOrigenDeDatos = "AmaproAccess" Then
                        vtexto = "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                    Else
                        vtexto = "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                    End If
                    vtexto = vtexto & UCase(Txttexto.Item(6).Text) & "', '" 'EMPLEADO
                    vtexto = vtexto & UCase(Txttexto.Item(1).Text) & "', " 'MANTENIMIENTO
                    vtexto = vtexto & Txttexto.Item(2).Text & ", '" 'MINUTOS
                    vtexto = vtexto & Txttexto.Item(7).Text & "', '" 'DESCRIPCION
                    vtexto = vtexto & GUsuario & "', "
                    vtexto = vtexto & VDocumento & ", '"
                    vtexto = vtexto & GEmpresa & "', "
                    vtexto = vtexto & "To_Date('" & Date & " " & Format(MskParIni.Text, "HH:mm") & "', 'dd/mm/yyyy hh24:mi'), "
                    vtexto = vtexto & "To_Date('" & Date & " " & Format(MskParFin.Text, "HH:mm") & "', 'dd/mm/yyyy hh24:mi')"
                    
                    'VALUES(To_Date('" & TxtFec.Text & " " & Format(Now.ToLocalTime, "HH:mm") & "', 'dd/mm/yyyy hh24:mi'),
                    
                    Conexion.Execute "Insert Into M_ActividadesDiarias Values(" & vtexto & ")"
                    
                    If Err <> 0 Then
                        MousePointer = 0
                        Conexion.RollbackTrans
                        MsgBox "No Se Grabo La Actividad " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                        Err.Clear
                        Exit Sub
                    End If
                    
                    Cont = 1
                    
                        
        'TERMINA LA TRANSACCION
        Conexion.CommitTrans
            
        MousePointer = 0
    
            

    Txttexto.Item(1).Text = ""
    Txttexto.Item(2).Text = ""
    Txttexto.Item(7).Text = ""
    Txttexto.Item(6).SetFocus
    
    
    
End Sub


Private Sub CmdSaleBusqueda_Click()
    FrameBusqueda.Visible = False
End Sub




Private Sub DbGridBusqueda_DblClick()
On Error Resume Next
        If BMantenimiento = True Then
                Txttexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(1).SetFocus
        ElseIf BEmpleado = True Then
                Txttexto.Item(6).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(6).SetFocus
        End If
        If Err <> 0 Then
        End If
        
        FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DbGridBusqueda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 43 Then
    On Error Resume Next
        If BMantenimiento = True Then
                Txttexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(1).SetFocus
        ElseIf BEmpleado = True Then
                Txttexto.Item(6).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(6).SetFocus
        End If
        If Err <> 0 Then
        End If
        FrameBusqueda.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF2 Then
                CmdGrabar_Click
            End If
End Sub

Private Sub Form_Load()
    MskFec.Text = Format(Date, "dd/mm/yyyy")
    Txttexto.Item(2).Text = "0"
        
    'HABILITA EL GRID PARA QUE PUEDAN CAMBIAR EL TAMAÑO DE LAS COLUMNAS
'    MSFlexGrid1.AllowUserResizing = flexResizeBoth
 '   MSFlexGrid1.AllowBigSelection = True
    
    'ANCHO DE COLUMNAS DEL GRID DE DETALLE
  '  anchocolumnasdetalle
    
    'ENCABEZADO DE CADA COLUMNA PARA EL GRID DE DETALLE
   ' EncabezadoDetalle
    
    
End Sub

Sub anchocolumnasdetalle()
    'MSFlexGrid1.ColWidth(0) = 10
    'MSFlexGrid1.ColWidth(1) = 2500
    'MSFlexGrid1.ColWidth(2) = 4600
End Sub
Sub EncabezadoDetalle()
    
    'MSFlexGrid1.Col = 1
    'MSFlexGrid1.CellBackColor = "&H80000004"
    'MSFlexGrid1.Text = "Codigo"
    'MSFlexGrid1.Col = 2
    'MSFlexGrid1.CellBackColor = "&H80000004"
    'MSFlexGrid1.Text = "Descripcion"

End Sub

Private Sub MskFec_GotFocus()
        MskFec.SelStart = 0
        MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub MskFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub




Private Sub MskParFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskParFin_LostFocus()
On Error Resume Next
        VHoraInicial = MskParIni.Text
        VHoraFinal = MskParFin.Text
        
        Txttexto.Item(2).Text = DateDiff("n", VHoraInicial, VHoraFinal)
        
                                
        If Err <> 0 Then
            MsgBox Err.Description
        End If
End Sub

Private Sub MskParIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    TxtBusqueda.SetFocus
End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda2 = New ADODB.Recordset
            If BMantenimiento = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda2, "Select Codigo, Descripcion From M_TiposMantenimiento where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda2, "Select Codigo, Descripcion From M_TiposMantenimiento where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda2, "Select Codigo, Descripcion From M_TiposMantenimiento where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda2, "Select Codigo, Descripcion From M_TiposMantenimiento where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            ElseIf BEmpleado = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda2, "Select No_Emple, Nombre From arplme where Nombre Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda2, "Select No_Emple, Nombre From arplme where UPPER(Nombre) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda2, "Select No_Emple, Nombre From arplme where No_Emple Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda2, "Select No_Emple, Nombre From arplme where UPPER(No_Emple) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            End If
            
                    
                    Set DBGridBusqueda.DataSource = RBusqueda2
                    DBGridBusqueda.Columns(1).Width = "4000"
            
            
            
    
End Sub

Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub



Private Sub Txttexto_Change(Index As Integer)
        If Index = 1 Then
                Set RBuscaMantenimiento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaMantenimiento, "Select Descripcion From M_TiposMantenimiento Where Codigo = '" & Txttexto.Item(1).Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaMantenimiento, "Select Descripcion From M_TiposMantenimiento Where UPPER(Codigo) = '" & UCase(Txttexto.Item(1).Text) & "'")
                    End If
                    If RBuscaMantenimiento.RecordCount > 0 Then
                        LblMan.Caption = RBuscaMantenimiento!descripcion
                    Else
                        LblMan.Caption = ""
                    End If
        
        ElseIf Index = 6 Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where No_Emple = '" & Txttexto.Item(6).Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Nombre From arplme Where UPPER(No_Emple) = '" & UCase(Txttexto.Item(6).Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblEmpl.Caption = RBuscaEmpleado!Nombre
                    Else
                        LblEmpl.Caption = ""
                    End If
        End If

End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        Set RBusqueda = New ADODB.Recordset
            
        TxtBusqueda.Text = ""
        
        If Index = 1 Then
            BMantenimiento = True
            BEmpleado = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_TiposMantenimiento")
        ElseIf Index = 6 Then
            BMantenimiento = False
            BEmpleado = True
            Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From arplme")
        End If
                
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "5000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        Txttexto.Item(Index).SelStart = 0
        Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 7 Then
        Else
            SendKeys "{tab}"
        End If
    End If
    
    If KeyAscii = 43 Then
        Set RBusqueda = New ADODB.Recordset
            
        TxtBusqueda.Text = ""
        
        If Index = 1 Then
            BMantenimiento = True
            BEmpleado = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_TiposMantenimiento")
        ElseIf Index = 6 Then
            BMantenimiento = False
            BEmpleado = True
            Call Abrir_Recordset(RBusqueda, "Select No_Emple, Nombre From arplme")
        End If
                
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "5000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus
    
    End If

End Sub
