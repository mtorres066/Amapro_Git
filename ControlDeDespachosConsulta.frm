VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ControlDeDespachosConsulta 
   Caption         =   "Consulta De Control De Despachos"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ControlDeDespachosConsulta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Framebuscar 
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
      Height          =   8055
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6855
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   12091
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
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7920
         Picture         =   "ControlDeDespachosConsulta.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
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
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Tipo De Producto"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Codigo"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo De Fecha"
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
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton OptFec 
         Caption         =   "Fecha Arribo"
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
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton OptFec 
         Caption         =   "Fecha Despacho"
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
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo De Factura "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   20
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton OptFac 
         Caption         =   "Que No Han Venido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptFac 
         Caption         =   "Que Ya Vinieron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton OptFac 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DbGrid1 
      Height          =   6375
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin VB.CommandButton CmdImprimir 
      Height          =   495
      Left            =   10440
      Picture         =   "ControlDeDespachosConsulta.frx":293C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11160
      Picture         =   "ControlDeDespachosConsulta.frx":2A86
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   495
      Left            =   9720
      Picture         =   "ControlDeDespachosConsulta.frx":4AF8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Consultar"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Descripcion"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   55771139
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   7560
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   55771139
      CurrentDate     =   38212
   End
   Begin VB.Label Label4 
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
      Left            =   6960
      TabIndex        =   11
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label3 
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
      Left            =   6960
      TabIndex        =   10
      Top             =   240
      Width           =   555
   End
   Begin VB.Label LblDes 
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
      Left            =   7680
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label LblEti 
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
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "ControlDeDespachosConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RConsulta As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim BProveedor As Boolean
Dim BFicha As Boolean
Dim BCodigo As Boolean
Dim BTipo As Boolean

Dim VDia As String
Dim VMes As String
Dim VAño  As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String

Dim Criteria As String



Private Sub CmdConsultar_Click()
        Set RConsulta = New ADODB.Recordset
        'CAMPOS
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = "Select R.FechaDespacho, R.FechaArribo, P.Descripcion, R.Codigo, F.Descrip, R.Cantidad, R.Factura, R.MontoDolares, R.Piloto, R.Transportista from ProductoEnTransito R, Proveedores P, FichaTecnica F"
            Else 'ORACLE
                Criteria = "Select R.FechaDespacho, R.FechaArribo, P.Descripcion, R.Codigo, F.Descrip, R.Cantidad, R.Factura, R.MontoDolares, (R.MontoDolares/R.Cantidad*1000), R.Piloto, R.Transportista from ProductoEnTransito R, Proveedores P, FichaTecnica F"
            End If
        
        
        'FECHA DESPACHO
        If OptFec.Item(0).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " where R.FechaDespacho >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And R.FechaDespacho <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
            Else 'ORACLE
                Criteria = Criteria & " where R.FechaDespacho >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And R.FechaDespacho <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')"
            End If
        'FECHA ARRIBO
        ElseIf OptFec.Item(1).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " where R.FechaArribo >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And R.FechaArribo <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
            Else 'ORACLE
                Criteria = Criteria & " where R.FechaArribo >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And R.FechaArribo <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')"
            End If
        End If
        
        'PROVEEDOR
        If OptOpc.Item(0).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " And R.Proveedor Like '%" & TxtTexto.Text & "%' And R.Proveedor = P.CodigoProveedor And R.Codigo = F.Esp_Tec"
            Else 'ORACLE
                Criteria = Criteria & " And UPPER(R.Proveedor) = UPPER(P.CodigoProveedor) And UPPER(R.Codigo) = UPPER(F.Esp_Tec)"
            End If
        'DESCRIPCION
        ElseIf OptOpc.Item(1).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " And R.Codigo = F.Esp_Tec And F.Descrip Like '%" & TxtTexto.Text & "%' And R.Proveedor = P.CodigoProveedor"
            Else 'ORACLE
                Criteria = Criteria & " And UPPER(R.Codigo) = UPPER(F.Esp_Tec) And UPPER(F.Descrip) Like '%" & UCase(TxtTexto.Text) & "%' And UPPER(R.Proveedor) = UPPER(P.CodigoProveedor)"
            End If
        'CODIGO
        ElseIf OptOpc.Item(2).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " And R.Codigo Like '%" & TxtTexto.Text & "%' And R.Proveedor = P.CodigoProveedor And R.Codigo = F.Esp_Tec"
            Else 'ORACLE
                Criteria = Criteria & " And UPPER(R.Codigo) Like '%" & UCase(TxtTexto.Text) & "%' And UPPER(R.Proveedor) = UPPER(P.CodigoProveedor) And UPPER(R.Codigo) = UPPER(F.Esp_Tec)"
            End If
        'TIPO
        ElseIf OptOpc.Item(3).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Criteria = Criteria & " And R.Codigo = F.Esp_Tec And F.Tipo Like '%" & TxtTexto.Text & "%' And R.Proveedor = P.CodigoProveedor"
            Else 'ORACLE
                Criteria = Criteria & " And UPPER(R.Codigo) = UPPER(F.Esp_Tec) And UPPER(F.Tipo) Like '%" & UCase(TxtTexto.Text) & "%' And UPPER(R.Proveedor) = UPPER(P.CodigoProveedor)"
            End If
        End If
                        
        If GOrigenDeDatos = "AmaproAccess" Then
                    'NO HAN VENIDO
                    If OptFac.Item(0).Value = True Then
                        Criteria = Criteria & " And R.Recibida = false Order By R.FechaArribo Desc, R.Factura Desc"
                    'YA VINIERON
                    ElseIf OptFac.Item(1).Value = True Then
                        Criteria = Criteria & " And R.Recibida = true Order By R.FechaArribo Desc, R.Factura Desc"
                    'TODAS
                    ElseIf OptFac.Item(2).Value = True Then
                        Criteria = Criteria & " Order By R.FechaArribo Desc, R.Factura Desc"
                    End If
        Else
                    'NO HAN VENIDO
                    If OptFac.Item(0).Value = True Then
                        Criteria = Criteria & " And R.Recibida = 0 Order By R.FechaArribo Desc, R.Factura Desc"
                    'YA VINIERON
                    ElseIf OptFac.Item(1).Value = True Then
                        Criteria = Criteria & " And R.Recibida = -1 Order By R.FechaArribo Desc, R.Factura Desc"
                    'TODAS
                    ElseIf OptFac.Item(2).Value = True Then
                        Criteria = Criteria & " Order By R.FechaArribo Desc, R.Factura Desc"
                    End If
        End If
        
        Call Abrir_Recordset(RConsulta, Criteria)
        Set DbGrid1.DataSource = RConsulta
        
        DbGrid1.Columns(0).Width = "1000"
        DbGrid1.Columns(1).Width = "1000"
        DbGrid1.Columns(2).Width = "2000"
        DbGrid1.Columns(3).Width = "1000"
        DbGrid1.Columns(4).Width = "3000"
        DbGrid1.Columns(5).Width = "1000"
        DbGrid1.Columns(6).Width = "1000"
        DbGrid1.Columns(7).Width = "1000"
        DbGrid1.Columns(8).Width = "1000"
        DbGrid1.Columns(9).Width = "1000"
'        DbGrid1.Columns(10).Width = "1000"
        
        DbGrid1.Columns(5).NumberFormat = "#,###,##0.00"
        DbGrid1.Columns(7).NumberFormat = "#,###,##0.00"
        DbGrid1.Columns(8).NumberFormat = "#,###,##0.00"

        DbGrid1.Columns(5).Alignment = dbgRight
        DbGrid1.Columns(7).Alignment = dbgRight
        DbGrid1.Columns(8).Alignment = dbgRight
        
        DbGrid1.Columns(5).Caption = "Cantidad"
        DbGrid1.Columns(7).Caption = "Dolares"
        'DbGrid1.Columns(8).Caption = "Costo Millar"
        
End Sub

Private Sub CmdImprimir_Click()

                 VDia = Day(DtpFecIni.Value)
                 VMes = Month(DtpFecIni.Value)
                 VAño = Year(DtpFecIni.Value)
                 VDia2 = Day(DTPFecFin.Value)
                 VMes2 = Month(DTPFecFin.Value)
                 VAño2 = Year(DTPFecFin.Value)
                
                'DESPACHO
                If OptFec.Item(0).Value = True Then
                        GCriteriaReporte = "{ProductoEnTransito.FechaDespacho} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                        'GCriteriaReporte = "{ProductoEnTransito.FechaDespacho} >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And {ProductoEnTransito.FechaDespacho} <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
                        GTituloReporte = "Fechas De Despacho Del " & DtpFecIni.Value & " Al " & DTPFecFin.Value
                'ARRIBO
                ElseIf OptFec.Item(1).Value = True Then
                        GCriteriaReporte = "{ProductoEnTransito.FechaArribo} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                        'GCriteriaReporte = "{ProductoEnTransito.FechaArribo} >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And {ProductoEnTransito.Arribo} <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#"
                        GTituloReporte = "Fechas De Arribo Del " & DtpFecIni.Value & " Al " & DTPFecFin.Value
                End If
                
                
                'PROVEEDOR
                If OptOpc.Item(0).Value = True Then
                            GTituloReporte = GTituloReporte & " Proveedor " & TxtTexto.Text & " " & LblDes.Caption
                            GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({ProductoEnTransito.Proveedor}) Like '*" & UCase(TxtTexto.Text) & "*'"
                            'GCriteriaReporte = "{ProductoEnTransito.FechaDespacho} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProductoEnTransito.Proveedor} Like '*" & TxtTexto.Text & "*'"
                'DESCRIPCION
                ElseIf OptOpc.Item(1).Value = True Then
                            GTituloReporte = GTituloReporte & " Descripcion " & TxtTexto.Text & " " & LblDes.Caption
                            GCriteriaReporte = GCriteriaReporte & " and UPPERCASE({FichaTecnica.Descrip}) Like '*" & UCase(TxtTexto.Text) & "*'"
                            'GCriteriaReporte = "{ProductoEnTransito.FechaDespacho} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProductoEnTransito.Codigo} Like '*" & TxtTexto.Text & "*'"
                'CODIGO
                ElseIf OptOpc.Item(2).Value = True Then
                            GTituloReporte = GTituloReporte & " Codigo " & TxtTexto.Text & " " & LblDes.Caption
                            GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({ProductoEnTransito.Codigo}) Like '*" & UCase(TxtTexto.Text) & "*'"
                            'GCriteriaReporte = "{ProductoEnTransito.FechaArribo} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProductoEnTransito.Proveedor} Like '*" & TxtTexto.Text & "*'"
                'TIPO
                ElseIf OptOpc.Item(3).Value = True Then
                            GTituloReporte = GTituloReporte & " Tipo " & TxtTexto.Text & " " & LblDes.Caption
                            GCriteriaReporte = GCriteriaReporte & " and uppercase({FichaTecnica.Tipo}) Like '*" & UCase(TxtTexto.Text) & "*'"
                            'GCriteriaReporte = "{ProductoEnTransito.FechaArribo} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") and {FichaTecnica.Descrip} Like '*" & TxtTexto.Text & "*'"
                End If
                
                'NO HAN VENIDO
                If OptFac.Item(0).Value = True Then
                    GCriteriaReporte = GCriteriaReporte & " And {ProductoEnTransito.Recibida} = 0"
                    GTituloReporte = GTituloReporte & " Que No Han Venido " & LblDes.Caption
                'YA VINIERON
                ElseIf OptFac.Item(1).Value = True Then
                    GCriteriaReporte = GCriteriaReporte & " And {ProductoEnTransito.Recibida} = -1"
                    GTituloReporte = GTituloReporte & " Que Ya Vinieron " & LblDes.Caption
                'TODAS
                ElseIf OptFac.Item(2).Value = True Then
                    GTituloReporte = GTituloReporte & " Todas " & LblDes.Caption
                End If
        
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "ProductoEnTransito.rpt"
                Else
                    GNombreReporte = "ProductoEnTransitoO.rpt"
                End If
                                
                FrmReporte.Show
                
                
End Sub

Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub



Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
            RConsulta.Sort = RConsulta.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_DblClick()
    If BProveedor = True Then
        TxtTexto.Text = DBGridBusqueda.Columns(0)
        TxtTexto.SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Text = DBGridBusqueda.Columns(1)
        TxtTexto.SetFocus
    ElseIf BCodigo = True Then
        TxtTexto.Text = DBGridBusqueda.Columns(0)
        TxtTexto.SetFocus
    ElseIf BTipo = True Then
        TxtTexto.Text = DBGridBusqueda.Columns(0)
        TxtTexto.SetFocus
    End If
        FrameBuscar.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
            
        If KeyAscii = 43 Then
            If BProveedor = True Then
                TxtTexto.Text = DBGridBusqueda.Columns(0)
                TxtTexto.SetFocus
            ElseIf BFicha = True Then
                TxtTexto.Text = DBGridBusqueda.Columns(1)
                TxtTexto.SetFocus
            ElseIf BCodigo = True Then
                TxtTexto.Text = DBGridBusqueda.Columns(0)
                TxtTexto.SetFocus
            ElseIf BTipo = True Then
                TxtTexto.Text = DBGridBusqueda.Columns(0)
                TxtTexto.SetFocus
            End If
                FrameBuscar.Visible = False
        End If
End Sub

Private Sub Form_Load()
        
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
        
        CmdConsultar_Click
End Sub

Private Sub Form_Resize()
        
        DbGrid1.Height = Me.Height - 2500
        DbGrid1.Width = Me.Width - 500
End Sub

Private Sub OptOpc_Click(Index As Integer)
    If Index = 0 Then
        LblEti.Caption = "Proveedor"
    ElseIf Index = 1 Then
        LblEti.Caption = "Descripcion"
    ElseIf Index = 2 Then
        LblEti.Caption = "Codigo"
    ElseIf Index = 3 Then
        LblEti.Caption = "Tipo"
    End If
        TxtTexto.SetFocus
End Sub


Private Sub Txtbusqueda_Change()
                    Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                        If BProveedor = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BFicha = True Or BCodigo = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BTipo = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        End If
                    'OPCION DE CODIGO
                    Else
                        If BProveedor = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BFicha = True Or BCodigo = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BTipo = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos Where CodigoTipo Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos Where UPPER(CodigoTipo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        End If
                    End If
                            
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtTexto_Change()
        'PROVEEDORES
        If OptOpc.Item(0).Value = True Then
            Set RBuscaProveedor = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtTexto.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtTexto.Text) & "'")
                End If
                If RBuscaProveedor.RecordCount > 0 Then
                    LblDes.Caption = RBuscaProveedor!Descripcion
                Else
                    LblDes.Caption = ""
                End If
        'FICHA TECNICA
        ElseIf OptOpc.Item(1).Value = True Or OptOpc.Item(2).Value = True Then
            Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Text) & "'")
                End If
                If RBuscaFicha.RecordCount > 0 Then
                    LblDes.Caption = RBuscaFicha!Descrip
                Else
                    LblDes.Caption = ""
                End If
        'TIPO
        ElseIf OptOpc.Item(3).Value = True Then
            Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtTexto.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtTexto.Text) & "'")
                End If
                If RBuscaFicha.RecordCount > 0 Then
                    LblDes.Caption = RBuscaFicha!Descripcion
                Else
                    LblDes.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick()
            Set RBusqueda = New ADODB.Recordset
            If OptOpc.Item(0).Value = True Then
                BProveedor = True
                BFicha = False
                BCodigo = False
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf OptOpc.Item(1).Value = True Then
                BProveedor = False
                BFicha = True
                BCodigo = False
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf OptOpc.Item(2).Value = True Then
                BProveedor = False
                BFicha = False
                BCodigo = True
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf OptOpc.Item(3).Value = True Then
                BProveedor = False
                BFicha = False
                BCodigo = False
                BTipo = True
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos")
            End If
                Set DBGridBusqueda.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DBGridBusqueda.Columns(1).Width = "4000"
            

End Sub

Private Sub TxtTexto_GotFocus()
        TxtTexto.SelStart = 0
        TxtTexto.SelLength = Len(TxtTexto.Text)
End Sub

Private Sub TxtTexto_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
            If OptOpc.Item(0).Value = True Then
                BProveedor = True
                BFicha = False
                BCodigo = False
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf OptOpc.Item(1).Value = True Then
                BProveedor = False
                BFicha = True
                BCodigo = False
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf OptOpc.Item(2).Value = True Then
                BProveedor = False
                BFicha = False
                BCodigo = True
                BTipo = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf OptOpc.Item(3).Value = True Then
                BProveedor = False
                BFicha = False
                BCodigo = False
                BTipo = True
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion from FichaTecnicaTipos")
            End If
            
        End If
    
End Sub
