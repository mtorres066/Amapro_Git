VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraVentas 
   Caption         =   "Genera Ventas"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "GeneraVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
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
      Height          =   6135
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
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
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7560
         Picture         =   "GeneraVentas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4092
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGridVentas 
      Height          =   4215
      Left            =   120
      TabIndex        =   28
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
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
   Begin VB.TextBox TxtBodDiaUni 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox TxtBod3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox TxtBodDiaDol 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox TxtCom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton CmdGenerar 
      Height          =   615
      Left            =   10440
      Picture         =   "GeneraVentas.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Generar"
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65536003
      CurrentDate     =   38023
   End
   Begin VB.TextBox TxtBod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65536003
      CurrentDate     =   38023
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   615
      Left            =   11160
      Picture         =   "GeneraVentas.frx":43EE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   615
      Left            =   9720
      Picture         =   "GeneraVentas.frx":6460
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Consultar Datos"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label LblBod5 
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
      Left            =   5640
      TabIndex        =   27
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bodega En Unidades Diarias"
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
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   26
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label LblBod3 
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
      Left            =   5640
      TabIndex        =   25
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bodega Unidades Acumulados"
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
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   24
      Top             =   480
      Width           =   2595
   End
   Begin VB.Label LblBod2 
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
      Left            =   5640
      TabIndex        =   17
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bodega En Dolares Diarios"
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
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   16
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label LblCom 
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
      Left            =   5640
      TabIndex        =   15
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ventas De Compañia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   510
   End
   Begin VB.Label LblBod 
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
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bodega Dolares Acumulados"
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
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
Attribute VB_Name = "GeneraVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaBodega2 As New ADODB.Recordset
Dim RBuscaBodega3 As New ADODB.Recordset
Dim RBuscaBodega5 As New ADODB.Recordset
Dim RAgregaVentas As New ADODB.Recordset
Dim RAgregaVentasDetalle As New ADODB.Recordset
Dim RAgruparVentasDetalle As New ADODB.Recordset
Dim RAgruparVentasDetalle2 As New ADODB.Recordset

Dim RBuscaCompañia As ADODB.Recordset
Dim RVentas As ADODB.Recordset
Dim RVentas2 As ADODB.Recordset
Dim RVerVentas As ADODB.Recordset
Dim RBusqueda As ADODB.Recordset

Dim BBodega1 As Boolean
Dim BBodega2 As Boolean
Dim BBodega3 As Boolean
Dim BBodega4 As Boolean

Dim VTexto As String

Private Sub CmdConsultar_Click()
On Error Resume Next
MousePointer = 11
        
           'INICIALIZA EL RECORDSET
           Set RVerVentas = New ADODB.Recordset
           
           
           'ABRE EL RECORDSETE
           'BUSCA LAS VENTAS EN UN RANGO DE FECHAS Y DE UNA COMPAÑIA Y QUE EL ESTADO SEA DIFERENTE A PENDIENTE
           'Y QUE EL INDICADOR DE ANULADO SEA DIFERENTE A A
           Call Abrir_Recordset(RVerVentas, "Select E.Fecha_Operacion, E.Fecha, E.No_Factu, E.No_Fisico, D.Pedido, D.Total, E.Tipo_Cambio, D.No_Arti, (D.Total/E.Tipo_Cambio) From ArfaFE E, ArfaFL D Where E.Fecha_Operacion >= TO_DATE('" & DtpFecIni.Value & "', 'DD/MM/YY') And E.Fecha_Operacion <= TO_DATE('" & DTPFecFin.Value & "', 'DD/MM/YY') And E.No_Cia = '" & TxtCom.Text & "' And E.Estado <> 'P' And Ind_Anu_Dev Is Null And E.No_Cia = D.No_Cia And E.No_Factu = D.No_Factu")
           'Call Abrir_Recordset(RVerVentas, "Select * From ArfaFl")
           'LLENA EL GRID CON EL RECORDSET
           Set DataGridVentas.DataSource = RVerVentas
           DataGridVentas.Columns(0).Width = "1000"
           DataGridVentas.Columns(1).Width = "1000"
           DataGridVentas.Columns(2).Width = "1000"
           DataGridVentas.Columns(3).Width = "1000"
           DataGridVentas.Columns(4).Width = "1000"
           DataGridVentas.Columns(5).Width = "1000"
           DataGridVentas.Columns(6).Width = "1000"
           DataGridVentas.Columns(7).Width = "1000"
           DataGridVentas.Columns(8).Width = "1000"
           
           
           If Err <> 0 Then
                MsgBox Err.Description & " " & Err.Number
           End If
           
           
MousePointer = 0

End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub


Private Sub DataGridVentas_HeadClick(ByVal ColIndex As Integer)
            RVerVentas.Sort = RVerVentas.Fields(ColIndex).Name
            'RVentas.Sort = RVentas.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BBodega1 = True Then
            TxtBod.Text = DBGridBusqueda.Columns(0).Text
            TxtBod.SetFocus
        ElseIf BBodega2 = True Then
            TxtBodDiaDol.Text = DBGridBusqueda.Columns(0).Text
            TxtBodDiaDol.SetFocus
        ElseIf BBodega3 = True Then
            TxtBod3.Text = DBGridBusqueda.Columns(0).Text
            TxtBod3.SetFocus
        ElseIf BBodega4 = True Then
            TxtBodDiaUni.Text = DBGridBusqueda.Columns(0).Text
            TxtBodDiaUni.SetFocus
        End If
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BBodega1 = True Then
                    TxtBod.Text = DBGridBusqueda.Columns(0).Text
                    TxtBod.SetFocus
                ElseIf BBodega2 = True Then
                    TxtBodDiaDol.Text = DBGridBusqueda.Columns(0).Text
                    TxtBodDiaDol.SetFocus
                ElseIf BBodega3 = True Then
                    TxtBod3.Text = DBGridBusqueda.Columns(0).Text
                    TxtBod3.SetFocus
                ElseIf BBodega4 = True Then
                    TxtBodDiaUni.Text = DBGridBusqueda.Columns(0).Text
                    TxtBodDiaUni.SetFocus
                End If
                    FrameBusqueda.Visible = False
        End If

End Sub

Private Sub Form_Load()
           
           DtpFecIni.Value = Date
           DTPFecFin.Value = Date
           
           
End Sub

Private Sub CmdGenerar_Click()
On Error Resume Next
MousePointer = 11
        Set RBuscaBodega = New ADODB.Recordset
        'REVISA LA BODEGA EN DOLARES
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod.Text) & "'")
            End If
            If RBuscaBodega.RecordCount > 0 Then
            Else
                MsgBox "Bodega En Dolares No Existe", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
                TxtBodDiaDol.SetFocus
                Exit Sub
            End If
        
        'REVISA LA BODEGA EN UNIDADES
        Set RBuscaBodega3 = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega3, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod3.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega3, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod3.Text) & "'")
            End If
            If RBuscaBodega3.RecordCount > 0 Then
            Else
                MsgBox "Bodega En Unidades No Existe", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
                TxtBod3.SetFocus
                Exit Sub
            End If


           'INICIALIZA EL RECORDSET
           
           'Set RAgregaVentas = Db.OpenRecordset("Select * From Ventas")
           
           'BORRA EL DETALLE DE VENTAS DEL RANGO DE OPERACION
           If GOrigenDeDatos = "AmaproAccess" Then
                Conexion.Execute "Delete From VentasDetalle Where FechaOperacion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaOperacion <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
           Else
                Conexion.Execute "Delete From VentasDetalle Where FechaOperacion >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaOperacion <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')"
           End If
           
                If Err.Number <> 0 Then
                    MsgBox "Error En Borrar Detalle Ventas" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If
           
           'BORRA EL RESUMEN DE VENTAS DEL RANGO DE OPERACION
           If GOrigenDeDatos = "AmaproAccess" Then
                Conexion.Execute "Delete From Ventas Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
           Else
                Conexion.Execute "Delete From Ventas Where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')"
           End If
           
                If Err.Number <> 0 Then
                    MsgBox "Error En Borrar Detalle Ventas" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If
           
           'INIICIALIZA EL recordset
           'Set RAgregaVentasDetalle = Db.OpenRecordset("Select * From VentasDetalle")
           
           'ABRE EL RECORDSETE
           'BUSCA LAS VENTAS EN UN RANGO DE FECHAS Y DE UNA COMPAÑIA Y QUE EL ESTADO SEA DIFERENTE A PENDIENTE
           'Y QUE EL INDICADOR DE ANULADO SEA DIFERENTE A A
           
           'VENTAS EN QUETZALES AL EXTERIOR_______________________________________________
           Set RVentas = New ADODB.Recordset
                    'Call Abrir_Recordset(RVentas, "Select E.Fecha_Operacion, E.Fecha, E.No_Factu, E.No_Fisico, D.Total, E.Tipo_Cambio, D.No_Arti, (D.Total/E.Tipo_Cambio), D.Pedido From ArfaFE E, ArfaFL D Where E.Fecha_Operacion >= TO_DATE('" & DtpFecIni.Value & "', 'DD/MM/YY') And E.Fecha_Operacion <= TO_DATE('" & DTPFecFin.Value & "', 'DD/MM/YY') And E.No_Cia = '" & TxtCom.Text & "' And E.Tipo_Doc <> '36' And E.Estado <> 'P' And Ind_Anu_Dev Is Null And E.No_Cia = D.No_Cia And E.No_Factu = D.No_Factu")
                    Call Abrir_Recordset(RVentas, "Select E.Fecha_Operacion, E.Fecha, E.No_Factu, E.No_Fisico, (D.Total-D.Imp_Incluido), E.Tipo_Cambio, D.No_Arti, ((D.Total-D.Imp_Incluido)/E.Tipo_Cambio), D.Pedido From ArfaFE E, ArfaFL D Where E.Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy') And E.Fecha <= TO_DATE('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(E.No_Cia) = '" & UCase(TxtCom.Text) & "' And UPPER(E.Moneda) = 'P' And UPPER(E.Estado) <> 'P' And Ind_Anu_Dev Is Null And E.No_Cia = D.No_Cia And E.No_Factu = D.No_Factu")
           If RVentas.RecordCount > 0 Then
                
                'INICIA LA TRANSACCION
                Conexion.BeginTrans
                        
                        Do Until RVentas.EOF
                                        VTexto = ""
                             'AGREGA EL DETALLE DE VENTAS A AMAPRO
                              'Db.Execute "Insert Into VentasDetalle (FechaOperacion, Fecha, NoFactura, NoFisico, TotalQuetzales, TazaCambio, FichaTecnica, TotalDolares, TotalUnidades, Usuario) VALUES(#" & RVentas(0) & "#, #" & RVentas(1) & "#, " & RVentas(2) & ", " & RVentas(3) & ", " & RVentas(4) & ", " & RVentas(5) & ", '" & RVentas(6) & "', " & RVentas(7) & ", " & RVentas(8) & ", '" & GUsuario & "')"
                                      
                                         If GOrigenDeDatos = "AmaproAccess" Then
                                              VTexto = "#" & Format(RVentas(0), "mm/dd/yyyy") & "#, " 'FECHA OPERACION
                                         Else 'ORACLE
                                              VTexto = "To_Date('" & RVentas(0) & "', 'dd/mm/yyyy')" & ", " 'FECHA OPERACION
                                         End If
                                         If GOrigenDeDatos = "AmaproAccess" Then
                                              VTexto = VTexto & "#" & Format(RVentas(1), "mm/dd/yyyy") & "#, "  'FECHA
                                         Else 'ORACLE
                                              VTexto = VTexto & "To_Date('" & RVentas(1) & "', 'dd/mm/yyyy')" & ", " 'FECHA
                                         End If
                                         VTexto = VTexto & RVentas(2) & ", " 'NOFCTURA
                                         VTexto = VTexto & RVentas(3) & ", " 'NOFISICO
                                         VTexto = VTexto & RVentas(4) & ", " 'TOTAL QUETZALES
                                         VTexto = VTexto & RVentas(5) & ", '" 'TAZA CAMBIO
                                         VTexto = VTexto & RVentas(6) & "', " 'FICHA TECNICA
                                         VTexto = VTexto & RVentas(7) & ", " 'TOTAL DOLARES
                                         VTexto = VTexto & RVentas(8) & ", '" 'TOTAL UNIDADES
                                         VTexto = VTexto & GUsuario & "'"  'HORAS HOMBRE
                                         'REALIZA EL INSERT
                                         Conexion.Execute "Insert Into VentasDetalle Values(" & VTexto & ")"
                             
                             If Err.Number <> 0 Then
                                 Conexion.RollbackTrans
                                 MousePointer = 0
                                 MsgBox "Error en agrega ventas detalle quetzales " & RVentas(6) & "     " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                 Err.Clear
                                 Exit Sub
                             End If
                             
                             'MUEVE AL SIGUIENTE REGISTRO
                             RVentas.MoveNext
                        Loop
           
                'TERMINA LA TRANSACCION
                Conexion.CommitTrans
          Else
                
          End If
           
           Set RVentas2 = New ADODB.Recordset
           'VENTAS EN DOLARES AL EXTERIOR_______________________________________________
           Call Abrir_Recordset(RVentas2, "Select E.Fecha_Operacion, E.Fecha, E.No_Factu, E.No_Fisico, ((D.Total-D.Imp_Incluido)*E.Tipo_Cambio), E.Tipo_Cambio, D.No_Arti, (D.Total-D.Imp_Incluido), D.Pedido From ArfaFE E, ArfaFL D Where E.Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'DD/MM/YY') And E.Fecha <= TO_DATE('" & DTPFecFin.Value & "', 'DD/MM/YY') And UPPER(E.No_Cia) = '" & UCase(TxtCom.Text) & "' And UPPER(E.Moneda) = 'D' And UPPER(E.Estado) <> 'P' And Ind_Anu_Dev Is Null And E.No_Cia = D.No_Cia And E.No_Factu = D.No_Factu")
           'Call Abrir_Recordset(RVentas2, "Select E.Fecha_Operacion, E.Fecha, E.No_Factu, E.No_Fisico, (D.Total*E.Tipo_Cambio), E.Tipo_Cambio, D.No_Arti, D.Total, D.Pedido From ArfaFE E, ArfaFL D Where E.Fecha_Operacion >= TO_DATE('" & DtpFecIni.Value & "', 'DD/MM/YY') And E.Fecha_Operacion <= TO_DATE('" & DTPFecFin.Value & "', 'DD/MM/YY') And E.No_Cia = '" & TxtCom.Text & "' And E.Tipo_Doc = '36' And E.Estado <> 'P' And Ind_Anu_Dev Is Null And E.No_Cia = D.No_Cia And E.No_Factu = D.No_Factu")
           
           'INICIA LA TRANSACCION
           Conexion.BeginTrans
           
           Do Until RVentas2.EOF
                'AGREGA EL DETALLE DE VENTAS A AMAPRO
                 'Db.Execute "Insert Into VentasDetalle (FechaOperacion, Fecha, NoFactura, NoFisico, TotalQuetzales, TazaCambio, FichaTecnica, TotalDolares, TotalUnidades, Usuario) VALUES(#" & RVentas(0) & "#, #" & RVentas(1) & "#, " & RVentas(2) & ", " & RVentas(3) & ", " & RVentas(4) & ", " & RVentas(5) & ", '" & RVentas(6) & "', " & RVentas(7) & ", " & RVentas(8) & ", '" & GUsuario & "')"
                            VTexto = ""
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(RVentas2(0), "mm/dd/yyyy") & "#, " 'FECHA OPERACION
                            Else 'ORACLE
                                 VTexto = "To_Date('" & RVentas2(0) & "', 'dd/mm/yyyy')" & ", " 'FECHA OPERACION
                            End If
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(RVentas2(1), "mm/dd/yyyy") & "#, "  'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & RVentas2(1) & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & RVentas2(2) & ", " 'NOFCTURA
                            VTexto = VTexto & RVentas2(3) & ", " 'NOFISICO
                            VTexto = VTexto & RVentas2(4) & ", " 'TOTAL QUETZALES
                            VTexto = VTexto & RVentas2(5) & ", '" 'TAZA CAMBIO
                            VTexto = VTexto & RVentas2(6) & "', " 'FICHA TECNICA
                            VTexto = VTexto & RVentas2(7) & ", " 'TOTAL DOLARES
                            VTexto = VTexto & RVentas2(8) & ", '" 'TOTAL UNIDADES
                            VTexto = VTexto & GUsuario & "'"  'HORAS HOMBRE
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into VentasDetalle Values(" & VTexto & ")"
                
                            
                If Err.Number <> 0 Then
                    Conexion.RollbackTrans
                    MousePointer = 0
                    MsgBox "Error en agrega ventas detalle en dolares" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                    Exit Sub
                End If
                
                'MUEVE AL SIGUIENTE REGISTRO
                RVentas2.MoveNext
           Loop
           
           'TERMIAN LA TRANSACCION
           Conexion.CommitTrans
           
           'VENTAS ACUMULADAS________________________________________________________________________________________
           
           'SUMA TODAS LAS VENTAS POR FICHA TECNICA DE ACUERDO AL RANGO DE FECHAS PARA IR SACANDO EL ACUMULADO
           Set RAgruparVentasDetalle = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAgruparVentasDetalle, "Select FichaTecnica, Sum(TotalDolares), Sum(TotalQuetzales), Sum(TotalUnidades) From VentasDetalle Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Group By Fichatecnica")
                Else
                    Call Abrir_Recordset(RAgruparVentasDetalle, "Select FichaTecnica, Sum(TotalDolares), Sum(TotalQuetzales), Sum(TotalUnidades) From VentasDetalle Where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Group By Fichatecnica")
                End If
                

           
            'INICIA LA TRANSACCION
            Conexion.BeginTrans
            
                'AGREGA EL TODOS LOS DATOS DE VENTAS
                Do Until RAgruparVentasDetalle.EOF
                    
                    'AGREGA TOTAL A BODEGA EN DOLARES_________________________________________________________
                    'Db.Execute "Insert Into Ventas (Fecha, FichaTecnica, Bodega, Cantidad, Usuario) VALUES(#" & DtpFecFin.Value & "#, '" & RAgruparVentasDetalle(0) & "', '" & TxtBod.Text & "', " & RAgruparVentasDetalle(1) & ", '" & GUsuario & "')"
                            VTexto = ""
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#, '" 'FECHA OPERACION
                            Else 'ORACLE
                                 VTexto = "To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & ", '" 'FECHA OPERACION
                            End If
                            VTexto = VTexto & RAgruparVentasDetalle(0) & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtBod.Text & "', " 'BODEGA
                            VTexto = VTexto & RAgruparVentasDetalle(1) & ", '" 'CANTIDAD
                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Ventas Values(" & VTexto & ")"
                    
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MousePointer = 0
                                MsgBox "Error en bodega dolares" & " " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                                Exit Sub
                            End If
                    
                    
                    
                            VTexto = ""
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#, '" 'FECHA OPERACION
                            Else 'ORACLE
                                 VTexto = "To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & ", '" 'FECHA OPERACION
                            End If
                            VTexto = VTexto & RAgruparVentasDetalle(0) & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtBod3.Text & "', " 'BODEGA
                            VTexto = VTexto & RAgruparVentasDetalle(3) & ", '" 'CANTIDAD
                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Ventas Values(" & VTexto & ")"
                    
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MousePointer = 0
                                MsgBox "Error en bodega Unidades" & " " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                                Exit Sub
                            End If
                    
                    'MUEVE AL SIGUIENTE REGISTRO
                    RAgruparVentasDetalle.MoveNext
                  Loop
                  
            'TERMINA LA CONNEXION
            Conexion.CommitTrans
           
           
           'VENTAS DIARIAS ------------------------------------------------------------------------------------------
           
           
           'SUMA TODAS LAS VENTAS POR FICHA TECNICA DE ACUERDO AL RANGO DE FECHAS PARA IR SACANDO EL ACUMULADO
           Set RAgruparVentasDetalle2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAgruparVentasDetalle2, "Select FichaTecnica, Sum(TotalDolares), Sum(TotalQuetzales), Sum(TotalUnidades) From VentasDetalle Where FechaOperacion = #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Group By Fichatecnica")
                Else
                    Call Abrir_Recordset(RAgruparVentasDetalle2, "Select FichaTecnica, Sum(TotalDolares), Sum(TotalQuetzales), Sum(TotalUnidades) From VentasDetalle Where FechaOperacion = To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Group By Fichatecnica")
                End If
                                               
                
           'INCIA LA TRANSACCION
           Conexion.BeginTrans
           
           'AGREGA EL TODOS LOS DATOS DE VENTAS
           Do Until RAgruparVentasDetalle2.EOF
                    
                    'AGREGA TOTAL A BODEGA EN DOLARES DIARIAS
                    'Db.Execute "Insert Into Ventas (Fecha, FichaTecnica, Bodega, Cantidad, Usuario) VALUES(#" & DtpFecFin.Value & "#, '" & RAgruparVentasDetalle2(0) & "', '" & TxtBodDiaDol.Text & "', " & RAgruparVentasDetalle2(1) & ", '" & GUsuario & "')"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#, '" 'FECHA OPERACION
                            Else 'ORACLE
                                 VTexto = "To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & ", '" 'FECHA OPERACION
                            End If
                            VTexto = VTexto & RAgruparVentasDetalle2(0) & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtBodDiaDol.Text & "', " 'BODEGA
                            VTexto = VTexto & RAgruparVentasDetalle2(1) & ", '" 'CANTIDAD
                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Ventas Values(" & VTexto & ")"
                    
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MousePointer = 0
                                MsgBox "Error en bodega dolares Unidades" & " " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                            End If
                            
                    
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#, '" 'FECHA OPERACION
                            Else 'ORACLE
                                 VTexto = "To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & ", '" 'FECHA OPERACION
                            End If
                            VTexto = VTexto & RAgruparVentasDetalle2(0) & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtBodDiaUni.Text & "', " 'BODEGA
                            VTexto = VTexto & RAgruparVentasDetalle2(3) & ", '" 'CANTIDAD
                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Ventas Values(" & VTexto & ")"
                    
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MousePointer = 0
                                MsgBox "Error en bodega Unidades" & " " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                            End If
                    
                'MUEVE AL SIGUIENTE REGISTRO
                RAgruparVentasDetalle2.MoveNext
           Loop
           
           'TERMINA LA TRANSACCION
           Conexion.CommitTrans
           
           
           MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
           
           'LIBERA EL RECORDSET DE MEMORIA
            'RVentas.Close
            'Set RVentas = Nothing
           
MousePointer = 0
                      
End Sub


Private Sub TxtBod_Change()
        'BUSCA BODEGA
        Set RBuscaBodega = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod.Text) & "'")
            End If
            If RBuscaBodega.RecordCount > 0 Then
                LblBod.Caption = RBuscaBodega!Descripcion
            Else
                LblBod.Caption = ""
            End If
End Sub

Private Sub TxtBod_DblClick()
                    BBodega1 = True
                    BBodega2 = False
                    BBodega3 = False
                    BBodega4 = False
                    Set RBusqueda = New ADODB.Recordset
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "4000"
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus

End Sub

Private Sub TxtBod_GotFocus()
            TxtBod.SelStart = 0
            TxtBod.SelLength = Len(TxtBod.Text)
End Sub

Private Sub TxtBod_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                
                If KeyAscii = 43 Then
                    BBodega1 = True
                    BBodega2 = False
                    BBodega3 = False
                    BBodega4 = False
                    Set RBusqueda = New ADODB.Recordset
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "4000"
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
                End If
End Sub

Private Sub txtboddiadol_Change()
        'BUSCA BODEGA
        Set RBuscaBodega2 = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega2, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBodDiaDol.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega2, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodDiaDol.Text) & "'")
            End If
            If RBuscaBodega2.RecordCount > 0 Then
                LblBod2.Caption = RBuscaBodega2!Descripcion
            Else
                LblBod2.Caption = ""
            End If

End Sub

Private Sub txtboddiadol_DblClick()
                    BBodega1 = False
                    BBodega2 = True
                    BBodega3 = False
                    BBodega4 = False
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
End Sub

Private Sub txtboddiadol_GotFocus()
                TxtBodDiaDol.SelStart = 0
                TxtBodDiaDol.SelLength = Len(TxtBodDiaDol.Text)
End Sub

Private Sub txtboddiadol_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                
                If KeyAscii = 43 Then
                    BBodega1 = False
                    BBodega2 = True
                    BBodega3 = False
                    BBodega4 = False
                    Set RBusqueda = New ADODB.Recordset
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
                End If
End Sub

Private Sub TxtBod3_Change()
        'BUSCA BODEGA
        Set RBuscaBodega3 = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega3, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod3.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega3, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod3.Text) & "'")
            End If
            If RBuscaBodega3.RecordCount > 0 Then
                LblBod3.Caption = RBuscaBodega3!Descripcion
            Else
                LblBod3.Caption = ""
            End If

End Sub

Private Sub TxtBod3_DblClick()
            BBodega1 = False
            BBodega2 = False
            BBodega3 = True
            BBodega4 = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
End Sub

Private Sub TxtBod3_GotFocus()
        TxtBod3.SelStart = 0
        TxtBod3.SelLength = Len(TxtBod3.Text)
End Sub

Private Sub TxtBod3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        
        If KeyAscii = 43 Then
            BBodega1 = False
            BBodega2 = False
            BBodega3 = True
            BBodega4 = False
            Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
        End If
End Sub


Private Sub TxtBodDiaUni_Change()
        'BUSCA BODEGA
        Set RBuscaBodega5 = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega5, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBodDiaUni.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega5, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodDiaUni.Text) & "'")
            End If
            If RBuscaBodega5.RecordCount > 0 Then
                LblBod5.Caption = RBuscaBodega5!Descripcion
            Else
                LblBod5.Caption = ""
            End If

End Sub

Private Sub TxtBodDiaUni_DblClick()
                    BBodega1 = False
                    BBodega2 = False
                    BBodega3 = False
                    BBodega4 = True
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus

End Sub

Private Sub TxtBodDiaUni_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                
                If KeyAscii = 43 Then
                    BBodega1 = False
                    BBodega2 = True
                    BBodega3 = False
                    BBodega4 = False
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
                End If

End Sub

Private Sub Txtbusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario where Codigo Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
                
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtCom_Change()
           'INICIALIZA EL RECORDSET
'           Set RBuscaCompañia = New ADODB.Recordset
           
           'ABRE EL RECORDSETE
'           Call Abrir_Recordset(RBuscaCompañia, "Select * From Tb_Companias")
           
'           LblCom.Caption = RBuscaCompañia!Descripcion

End Sub
