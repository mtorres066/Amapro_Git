VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ActividadesDiariasBorrar 
   BackColor       =   &H0080C0FF&
   Caption         =   "Borrar Actividades Diarias"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ActividadesDiariasBorrar.frx":0000
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
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   7215
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "ActividadesDiariasBorrar.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   7095
         Left            =   120
         TabIndex        =   12
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
   Begin VB.CommandButton CmdDen 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   9600
      Picture         =   "ActividadesDiariasBorrar.frx":3D6C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1095
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
      Picture         =   "ActividadesDiariasBorrar.frx":4334
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdVerDatos 
      Caption         =   "&Consultar"
      Height          =   855
      Left            =   8400
      Picture         =   "ActividadesDiariasBorrar.frx":484F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ver Datos"
      Top             =   120
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7095
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      GridColor       =   -2147483629
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "ActividadesDiariasBorrar"
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




Private Sub CmdDen_Click()
    On Error Resume Next
         
        
                        If IsNumeric(VDocumento) Then
                            
                        Else
                                MsgBox "No Ha Seleccionado Ninguna Actividad", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                        End If
        
        
        'INICIA LA TRANSACCION
        Conexion.BeginTrans
           
                    'COLOCA LA HORA EN QUE INICIO EL LOTE
                    Conexion.Execute "delete from M_ActividadesDiarias Where Documento = " & VDocumento
                    
                    
                    If Err <> 0 Then
                        MsgBox "Error Al Borrar Documento " & VDocumento & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                        Err.Clear
                        Conexion.RollbackTrans
                        Exit Sub
                    End If
                                        
                    
        'TERMINA LA TRANSACCION
        Conexion.CommitTrans
        
        verdatos
        MSHFlexGrid1_Click



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
    MSHFlexGrid1.Width = (Me.Width - 300)
    MSHFlexGrid1.Height = (Me.Height - 1700)
    
    
    If Err <> 0 Then
    End If
End Sub

Private Sub MSHFlexGrid1_Click()
On Error Resume Next
If MSHFlexGrid1.Text <> "" Then
    MSHFlexGrid1.Col = 0
    VDocumento = MSHFlexGrid1.Text
End If

End Sub

Private Sub MSHFlexGrid1_RowColChange()
    Static UltimaFila As Long 'Variable estatica, para almacenar LA ultima fila
    Dim Fila As Long, Columna As Long
    Dim FilaActual As Long, ColumnaActual As Long
    Dim strRuta As String
    
    'Cambio de color en la fila actual, segun el desplazamiento
    With MSHFlexGrid1
        .Redraw = False
        
        If .Row <> UltimaFila Then
            UltimaFila = .Row
            FilaActual = .Row
            ColumnaActual = .Col
            .Col = 0
            'Cambia el color de todo el MSHFlex
            For Fila = 1 To .Rows - 1
                .Row = Fila
                For Columna = 0 To .Cols - 1
                    .Col = Columna
                    If Fila = UltimaFila Then
                        'Todas las filas menos la actual
                        .CellBackColor = &HC0FFFF 'Blanco
                    Else
                        'La fila actual
                        .CellBackColor = &H8000000E 'Amarillo
                    End If
                Next
                'Regresa el alto de todas las filas a su valor original
                .RowHeight(Fila) = .RowHeight(0)
                'Quita la imagen de la ultima fila
                If .Col = .Cols - 1 Then Set .CellPicture = Nothing
                
                
            Next Fila
            'Regresa la fila y columna a sus valores originales
            .Row = FilaActual
            .Col = ColumnaActual
        End If
        'Se usa Redraw para evitar el parpadeo del control
        .Redraw = True
    End With

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
                    Set MSHFlexGrid1.DataSource = RVerDatos
                    
                            
                            MSHFlexGrid1.ColWidth(0) = "10" ' DOCUMENTO O LOTE
                            MSHFlexGrid1.ColWidth(1) = "1000" 'LINEA
                            MSHFlexGrid1.ColWidth(2) = "2000" 'PRODUCTO
                            MSHFlexGrid1.ColWidth(3) = "2000" 'PRODUCTO
                            MSHFlexGrid1.ColWidth(4) = "500" ' CAJAS
                            MSHFlexGrid1.ColWidth(5) = "5300" ' CAJAS
                            
                            MSHFlexGrid1.ColHeaderCaption(0, 0) = "Documento"
                            MSHFlexGrid1.ColHeaderCaption(0, 1) = "Fecha"
                            MSHFlexGrid1.ColHeaderCaption(0, 2) = "Empleado"
                            MSHFlexGrid1.ColHeaderCaption(0, 3) = "Mantenimiento"
                            MSHFlexGrid1.ColHeaderCaption(0, 4) = "Minutos"
                            MSHFlexGrid1.ColHeaderCaption(0, 5) = "Descripcion Del Trabajo"
                            
                            

                    
            MousePointer = 0
            
                If Err <> 0 Then
                    MsgBox "ERROR " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If

            

End Sub
