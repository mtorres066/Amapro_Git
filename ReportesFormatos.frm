VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesFormatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes De Formato"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "ReportesFormatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4335
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7646
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
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7320
         Picture         =   "ReportesFormatos.frx":2072
         Style           =   1  'Graphical
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtPro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   51970051
      CurrentDate     =   38147
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   51970051
      CurrentDate     =   38147
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   735
      Left            =   10440
      Picture         =   "ReportesFormatos.frx":40E4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1300
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Height          =   735
      Left            =   9000
      Picture         =   "ReportesFormatos.frx":45FF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1300
   End
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox TxtTransaccion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formatos De Registros De Inspeccion"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Cobros a Proveedores"
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   23
         Top             =   2280
         Width           =   2055
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Campanas Para Galon"
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   13
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Material De Empaque"
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Barniz En Polvo"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Barniz Liquido Y Sello Solve"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Alambre De Cobre"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Alambre Para Asas"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Anillos"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Fondos y Tapas"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Hojalata"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label LblProDes 
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
      Left            =   7560
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label LblPro 
      Caption         =   "Proveedor"
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
      Left            =   4920
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Index           =   2
      Left            =   7920
      TabIndex        =   24
      Top             =   3480
      Width           =   3660
   End
   Begin VB.Label LblFecFin 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label LblFecIni 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label LblTransaccion 
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
      Left            =   6120
      TabIndex        =   17
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label LblCodigo 
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
      Left            =   6120
      TabIndex        =   16
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   960
      Width           =   600
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Transaccion"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "ReportesFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaCodigo As New ADODB.Recordset
Dim RBuscaEntrada As New ADODB.Recordset
Dim RBuscaProveedor As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VDia As String
Dim VMes As String
Dim VAño As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String


Private Sub CmdImprimir_Click()
On Error Resume Next
                MousePointer = 11
                        
                'gtituloreporte = "Texto = ' Por Materia Prima " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                If OptOpcion.Item(9).Value = True Then
                Else
                        If Not IsNumeric(TxtTransaccion.Text) Then
                            MsgBox "La Transaccion Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                End If
                 
                'COBROS Descripcion
                If OptOpcion.Item(9).Value = True Then
                            VDia = Day(DtpFecIni.Value)
                            VMes = Month(DtpFecIni.Value)
                            VAño = Year(DtpFecIni.Value)
                            VDia2 = Day(DTPFecFin.Value)
                            VMes2 = Month(DTPFecFin.Value)
                            VAño2 = Year(DTPFecFin.Value)
                                 
                            GTituloReporte = "Desde " & DtpFecIni.Value & " Hasta " & DTPFecFin.Value
                            GCriteriaReporte = "{CobrosProveedor.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CobrosProveedor.Proveedor} Like '" & TxtPro.Text & "*'"
                'OTROS REPORTES
                Else
                            GCriteriaReporte = "{EncabezadoEntradasInventario.Documento} = " & TxtTransaccion.Text & " And UCASE({DetalleEntradasInventario.FichaTecnica}) Like '" & UCase(TxtCodigo.Text) & "*'"
                End If
                 
                If OptOpcion.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionHojalata.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionHojalataO.rpt"
                        End If
                ElseIf OptOpcion.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionFondos.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionFondosO.rpt"
                        End If
                ElseIf OptOpcion.Item(2).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionAnillos.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionAnillosO.rpt"
                        End If
                ElseIf OptOpcion.Item(3).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionAlambreAsas.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionAlambreAsasO.rpt"
                        End If
                ElseIf OptOpcion.Item(4).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionAlambreCobre.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionAlambreCobreO.rpt"
                        End If
                ElseIf OptOpcion.Item(5).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionBarnizLiquido.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionBarnizLiquidoO.rpt"
                        End If
                ElseIf OptOpcion.Item(6).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionBarnizPolvo.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionBarnizPolvoO.rpt"
                        End If
                ElseIf OptOpcion.Item(7).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionMaterialEmpaque.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionMaterialEmpaqueO.rpt"
                        End If
                ElseIf OptOpcion.Item(8).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FormatoInspeccionCampanas.rpt"
                        Else
                            GNombreReporte = "FormatoInspeccionCampanasO.rpt"
                        End If
                ElseIf OptOpcion.Item(9).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "CobrosProveedor.rpt"
                        Else
                            GNombreReporte = "CobrosProveedorO.rpt"
                        End If
                End If
                        
            


                MousePointer = 0
                FrmReporte.Show
                
                
                If Err > 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                End If

End Sub

Private Sub CmdSale_Click()
            FrameBuscar.Visible = False
End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        TxtPro.Text = DBGridBusqueda.Columns(0)
        TxtPro.SetFocus
        FrameBuscar.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            TxtPro.Text = DBGridBusqueda.Columns(0)
            TxtPro.SetFocus
            FrameBuscar.Visible = False
        End If
End Sub

Private Sub Form_Load()
            DtpFecIni.Value = Date
            DTPFecFin.Value = Date
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        'COBROS A PROVEEDORES
        If Index = 9 Then
            lblfecini.Visible = True
            LblFecFin.Visible = True
            DtpFecIni.Visible = True
            DTPFecFin.Visible = True
            LblPro.Visible = True
            LblProDes.Visible = True
            TxtPro.Visible = True
        Else
            lblfecini.Visible = False
            LblFecFin.Visible = False
            DtpFecIni.Visible = False
            DTPFecFin.Visible = False
            LblPro.Visible = False
            LblProDes.Visible = False
            TxtPro.Visible = False
        End If
End Sub

Private Sub TxtBusqueda_Change()
                    Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
                            
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "5000"

End Sub

Private Sub TxtCodigo_Change()
        Set RBuscaCodigo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaCodigo, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCodigo.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaCodigo, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodigo.Text) & "'")
            End If
            If RBuscaCodigo.RecordCount > 0 Then
                LblCodigo.Caption = RBuscaCodigo!Descrip
            Else
                LblCodigo.Caption = ""
            End If
End Sub

Private Sub TxtCodigo_GotFocus()
        TxtCodigo.SelStart = 0
        TxtCodigo.SelLength = Len(TxtCodigo.Text)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtPro_Change()
        Set RBuscaProveedor = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtPro.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtPro.Text) & "'")
            End If
            If RBuscaProveedor.RecordCount > 0 Then
                LblProDes.Caption = RBuscaProveedor!Descripcion
            Else
                LblProDes.Caption = ""
            End If
End Sub

Private Sub TxtPro_DblClick()
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                Set DBGridBusqueda.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtPro_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                
                If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    FrameBuscar.Visible = True
                    TxtBusqueda.SetFocus
                    DBGridBusqueda.Columns(1).Width = "4000"
                End If
End Sub

Private Sub TxtTransaccion_Change()
        If IsNumeric(TxtTransaccion.Text) Then
        Set RBuscaEntrada = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaEntrada, "Select E.NumeroDocumento, D.Descripcion From EncabezadoEntradasInventario E, Documentos D Where E.Documento = " & TxtTransaccion.Text & " And E.TipoDeDocumento = D.CodigoDocumento")
            Else
                Call Abrir_Recordset(RBuscaEntrada, "Select E.NumeroDocumento, D.Descripcion From EncabezadoEntradasInventario E, Documentos D Where E.Documento = " & TxtTransaccion.Text & " And UPPER(E.TipoDeDocumento) = UPPER(D.CodigoDocumento)")
            End If
            If RBuscaEntrada.RecordCount > 0 Then
                LblTransaccion.Caption = "Documento " & RBuscaEntrada(0) & Space(10) & RBuscaEntrada(1)
            Else
                LblTransaccion.Caption = ""
            End If
        End If
        
End Sub

Private Sub TxtTransaccion_GotFocus()
        TxtTransaccion.SelStart = 0
        TxtTransaccion.SelLength = Len(TxtTransaccion.Text)
End Sub

Private Sub TxtTransaccion_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
