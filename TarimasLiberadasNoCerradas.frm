VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TarimasLiberadasNoCerradas 
   Caption         =   "Tarimas Liberadas y No Cerradas En Inventario"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DbGridTarimas 
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13150
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
   Begin VB.CommandButton CmdSalida 
      Height          =   375
      Left            =   11280
      Picture         =   "TarimasLiberadasNoCerradas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   375
      Left            =   10680
      Picture         =   "TarimasLiberadasNoCerradas.frx":2072
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Consultar"
      Top             =   0
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPMes 
      Height          =   255
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MM/yyyy"
      Format          =   64880643
      CurrentDate     =   37578
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarimas Revisadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   7920
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarima Liberada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mes y Año"
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
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "TarimasLiberadasNoCerradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RTarimas As New ADODB.Recordset
Private Sub CmdConsultar_Click()
          'BUSCA TODAS LAS TARIMAS QUE ESTAN LIBERADAS Y NO SE HAN REBAJADO DEL INVENTARIO
           'DataTarimas.RecordSource = "SELECT DISTINCTROW P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas as P LEFT JOIN CierreBulto As DE ON P.Fec_PrdL = DE.FechaProduccion And P.LineaL = DE.Linea And P.Esp_TecL = DE.FichaTecnica And P.Tarima = DE.Tarima Where Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value)
            'DataTarimas.RecordSource = "SELECT P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.Esp_TecL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas as P, CierreBulto As DE Where P.Fec_PrdL not equal DE.FechaProduccion And P.LineaL not equal DE.Linea And P.Esp_TecL not equal DE.FichaTecnica And P.Tarima not equal DE.Tarima And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value)
            'DataTarimas.RecordSource = "SELECT DISTINCTROW P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas as P, CierreBulto As DE Where DE.FechaProduccion = P.Fec_PrdL And DE.Linea = P.LineaL And DE.FichaTecnica = P.Esp_TecL And DE.Tarima = P.Tarima And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value)
            'DataTarimas.RecordSource = "SELECT P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas as P, CierreBulto As DE Where P.Fec_PrdL = DE.FechaProduccion And P.LineaL = DE.Linea And P.Esp_TecL = DE.FichaTecnica And P.Tarima = DE.Tarima WHERE (((DE.FechaProduccion) Is Null)) And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value)
'  DataBatch.RecordSource = "SELECT DISTINCTROW P.Linea, P.Batch FROM Produccion as P LEFT JOIN DetalleEntradasProductoTermina As DE ON P.Linea = DE.Linea And P.Batch = DE.Batch Where (((DE.Batch) Is Null)) And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value) & " Group By P.Linea, P.Batch"

            Set RTarimas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "SELECT P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.Esp_TecL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas P LEFT JOIN CierreBulto C ON P.Fec_PrdL = C.FechaProduccion And P.LineaL = C.Linea And P.Esp_TecL = C.FichaTecnica And P.TarimaL = C.Tarima WHERE Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value) & " And (((C.Linea) Is Null) And ((C.FechaProduccion) is Null) And ((C.FichaTecnica) Is Null) And ((C.Tarima) Is Null)) And P.CalidadL <> 'C'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "SELECT P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.Fec_PrdL, P.LineaL, P.Esp_TecL, P.TarimaL, P.CalidadL, P.Revisados, P.NoConforme, P.Liberados, P.EnTarima FROM ProduccionLiberadaConTarimas P LEFT JOIN CierreBulto C ON P.Fec_PrdL = C.FechaProduccion And UPPER(P.LineaL) = UPPER(C.LineaProduccion) And UPPER(P.Esp_TecL) = UPPER(C.FichaTecnica) And P.TarimaL = C.Tarima WHERE To_Char(P.Fec_Prd, 'mm') = " & Month(DTPMes.Value) & " And To_Char(P.Fec_Prd, 'yyyy') = " & Year(DTPMes.Value) & " And (((UPPER(C.LineaProduccion)) Is Null) And ((C.FechaProduccion) is Null) And ((UPPER(C.FichaTecnica)) Is Null) And ((C.Tarima) Is Null)) And UPPER(P.CalidadL) <> 'C'")
                    'PARA VER LOS BULTOS QUE NO ESTAN EN INVENTARIO
                    'Call Abrir_Recordset(RTarimas, "SELECT P.Fec_Prd, P.Linea, P.Esp_Tec, P.Tarima, P.CodigoMateriaPrima, P.Bulto, P.FechaProduccion, P.LineaProduccion FROM ProduccionConMateriaPrima P LEFT JOIN DetalleEntradasInventario C ON P.FechaProduccion = C.FechaProduccion And UPPER(P.LineaProduccion) = UPPER(C.Linea) And UPPER(P.CodigoMateriaPrima) = UPPER(C.FichaTecnica) And P.Bulto = C.Tarima WHERE (((UPPER(C.Linea)) Is Null) And ((C.FechaProduccion) is Null) And ((UPPER(C.FichaTecnica)) Is Null) And ((C.Tarima) Is Null))")
                End If
           
           Set DbGridTarimas.DataSource = RTarimas
           
           DbGridTarimas.Columns(0).Width = "1000"
           DbGridTarimas.Columns(0).Caption = "Fecha"
           DbGridTarimas.Columns(1).Width = "400"
           DbGridTarimas.Columns(2).Width = "1400"
           DbGridTarimas.Columns(2).Caption = "Ficha Tecnica"
           DbGridTarimas.Columns(3).Width = "600"
           DbGridTarimas.Columns(4).Width = "1000"
           DbGridTarimas.Columns(4).Caption = "Fecha"
           DbGridTarimas.Columns(5).Width = "400"
           DbGridTarimas.Columns(6).Width = "1400"
           DbGridTarimas.Columns(6).Caption = "Ficha Tecnica"
           DbGridTarimas.Columns(7).Width = "600"
           DbGridTarimas.Columns(8).Width = "300"
           DbGridTarimas.Columns(9).Width = "800"
           DbGridTarimas.Columns(10).Width = "800"
           DbGridTarimas.Columns(11).Width = "800"
           DbGridTarimas.Columns(12).Width = "800"
           DbGridTarimas.Columns(9).NumberFormat = "#,###,##0"
           DbGridTarimas.Columns(10).NumberFormat = "#,###,##0"
           DbGridTarimas.Columns(11).NumberFormat = "#,###,##0"
           DbGridTarimas.Columns(12).NumberFormat = "#,###,##0"
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DbGridTarimas_HeadClick(ByVal ColIndex As Integer)
        RTarimas.Sort = RTarimas.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
    DTPMes.Value = Date
End Sub
