VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ConsultaInventarioProductoTerminado 
   BackColor       =   &H00FF8080&
   Caption         =   "Consulta De Inventario Producto Terminado Y Materia Prima"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "ConsultaInventarioProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9900
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
      Height          =   8415
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   9855
      Begin MSDataGridLib.DataGrid Dbgridbusqueda 
         Height          =   7215
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   12726
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
         Left            =   9120
         Picture         =   "ConsultaInventarioProductoTerminado.frx":628A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo De Inventario"
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
      Height          =   1335
      Left            =   2640
      TabIndex        =   19
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton OptTipInv 
         BackColor       =   &H00FF8080&
         Caption         =   "Producto Terminado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptTipInv 
         BackColor       =   &H00FF8080&
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton OptTipInv 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid Dbgridfichatecnica 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   12303
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
   Begin VB.OptionButton OptFichaBodega 
      BackColor       =   &H00FF8080&
      Caption         =   "Bodega Y Ficha Tecnica"
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
      TabIndex        =   17
      Top             =   840
      Width           =   2535
   End
   Begin VB.OptionButton OptBodega 
      BackColor       =   &H00FF8080&
      Caption         =   "Bodega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton OptFichaTecnica 
      BackColor       =   &H00FF8080&
      Caption         =   "Ficha Tecnica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox TxtBodega 
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
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   615
      Left            =   9120
      Picture         =   "ConsultaInventarioProductoTerminado.frx":82FC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   720
      Width           =   735
   End
   Begin MSMask.MaskEdBox MskTotalEnvases 
      Height          =   285
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTotalTarimas 
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Tarimas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   675
   End
   Begin VB.Label LblBodega 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
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
      Left            =   4920
      TabIndex        =   8
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label LblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
   Begin MSForms.CommandButton CmdGenera 
      Default         =   -1  'True
      Height          =   615
      Left            =   8280
      TabIndex        =   1
      ToolTipText     =   "Generar Datos"
      Top             =   720
      Width           =   735
      PicturePosition =   327683
      Size            =   "1296;1085"
      Picture         =   "ConsultaInventarioProductoTerminado.frx":8817
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "ConsultaInventarioProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RInventario As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean

Dim VTexto As String



Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11
            Set RInventario = New ADODB.Recordset
            
            If OptFichaTecnica.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, Count(De.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE, FichaTecnica F Where DE.FichaTecnica = F.ESP_TEC And DE.Saldo > 0"
                    Else 'ORACLE
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, Count(De.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE, FichaTecnica F Where UPPER(DE.FichaTecnica) = UPPER(F.ESP_TEC) And DE.Saldo > 0"
                    End If
            ElseIf OptBodega.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE INNER JOIN FichaTecnica F ON DE.FichaTecnica = F.ESP_TEC Where De.Saldo > 0 AND DE.Bodega = '" & TxtBodega.Text & "'"
                    Else 'ORACLE
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE INNER JOIN FichaTecnica F ON DE.FichaTecnica = F.ESP_TEC Where De.Saldo > 0 AND UPPER(DE.Bodega) = '" & UCase(TxtBodega.Text) & "'"
                    End If
            ElseIf OptFichaBodega.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, De.Bodega, B.Descripcion, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE INNER JOIN FichaTecnica F ON DE.FichaTecnica = F.ESP_TEC, BodegasInventario B Where DE.Saldo > 0 AND DE.Bodega = B.CodigoBodega"
                    Else 'ORACLE
                        VTexto = "SELECT DE.FichaTecnica, F.Descrip, De.Bodega, B.Descripcion, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE INNER JOIN FichaTecnica F ON DE.FichaTecnica = F.ESP_TEC, BodegasInventario B Where DE.Saldo > 0 AND UPPER(DE.Bodega) = UPPER(B.CodigoBodega)"
                    End If
            End If
            
            'TIPO DE INVENTARIO _____________________________________________________________
            If OptTipInv.Item(0).Value = True Then
                                    
            ElseIf OptTipInv.Item(1).Value = True Then
                    VTexto = VTexto & " And F.TipoInventario = 'MATERIA PRIMA'"
            ElseIf OptTipInv.Item(2).Value = True Then
                    VTexto = VTexto & " And F.TipoInventario = 'PRODUCTO TERMINADO'"
            End If

            
            If OptFichaTecnica.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip"
                    Else 'ORACLE
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip"
                    End If
            ElseIf OptBodega.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip"
                    Else 'ORACLE
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip"
                    End If
            ElseIf OptFichaBodega.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip, DE.Bodega, B.Descripcion"
                    Else 'ORACLE
                        VTexto = VTexto & " Group By DE.FichaTecnica, F.Descrip, DE.Bodega, B.Descripcion"
                    End If
            End If
            
            Call Abrir_Recordset(RInventario, VTexto)
            
            Set Dbgridfichatecnica.DataSource = RInventario
            
            'POR BODEGA Y FICHA TECNICA
            If OptFichaBodega.Value = True Then
                    
                    Dbgridfichatecnica.Columns(0).Caption = "Ficha Tecnica"
                    Dbgridfichatecnica.Columns(1).Caption = "Descripcion"
                    Dbgridfichatecnica.Columns(2).Caption = "Bodega"
                    Dbgridfichatecnica.Columns(3).Caption = "Descripcion"
                    Dbgridfichatecnica.Columns(4).Caption = "Tarimas"
                    Dbgridfichatecnica.Columns(5).Caption = "Cantidad"
                    
                    Dbgridfichatecnica.Columns(0).Width = "1200"
                    Dbgridfichatecnica.Columns(1).Width = "4000"
                    Dbgridfichatecnica.Columns(2).Width = "400"
                    Dbgridfichatecnica.Columns(3).Width = "1500"
                    Dbgridfichatecnica.Columns(4).Width = "600"
                    Dbgridfichatecnica.Columns(5).Width = "1000"
                                              
                    Dbgridfichatecnica.Columns(4).NumberFormat = "#,###,##0"
                    Dbgridfichatecnica.Columns(5).NumberFormat = "#,###,##0.00"
                    Dbgridfichatecnica.Columns(4).Alignment = dbgRight
                    Dbgridfichatecnica.Columns(5).Alignment = dbgRight
                    
            Else
                    Dbgridfichatecnica.Columns(0).Caption = "Ficha Tecnica"
                    Dbgridfichatecnica.Columns(1).Caption = "Descripcion"
                    Dbgridfichatecnica.Columns(2).Caption = "Tarimas"
                    Dbgridfichatecnica.Columns(3).Caption = "Cantidad"
                    Dbgridfichatecnica.Columns(0).Width = "1300"
                    Dbgridfichatecnica.Columns(1).Width = "4500"
                    Dbgridfichatecnica.Columns(2).Width = "1200"
                    Dbgridfichatecnica.Columns(3).Width = "1200"
                    
                    Dbgridfichatecnica.Columns(2).NumberFormat = "#,###,##0"
                    Dbgridfichatecnica.Columns(3).NumberFormat = "#,###,##0.00"
                    Dbgridfichatecnica.Columns(2).Alignment = dbgRight
                    Dbgridfichatecnica.Columns(3).Alignment = dbgRight
                
            End If
            
            
            
            Set RTotal = New ADODB.Recordset
            
            If OptFichaTecnica.Value = True Then
                    Call Abrir_Recordset(RTotal, "SELECT Count(Saldo), Sum(Saldo) From DetalleEntradasInventario Where Saldo > 0")
            ElseIf OptBodega.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTotal, "SELECT Count(Saldo), Sum(Saldo) From DetalleEntradasInventario Where Saldo > 0 And Bodega = '" & TxtBodega.Text & "'")
                    Else 'oracle
                        Call Abrir_Recordset(RTotal, "SELECT Count(Saldo), Sum(Saldo) From DetalleEntradasInventario Where Saldo > 0 And UPPER(Bodega) = '" & UCase(TxtBodega.Text) & "'")
                    End If
            ElseIf OptFichaBodega.Value = True Then
                        Call Abrir_Recordset(RTotal, "SELECT Count(Saldo), Sum(Saldo) From DetalleEntradasInventario Where Saldo > 0 ")
            End If
                        
            If RTotal.RecordCount > 0 Then
                        If Not IsNull(RTotal(0)) Then
                            MskTotalTarimas.Text = RTotal(0)
                        Else
                            MskTotalTarimas.Text = "0"
                        End If
                        
                        If Not IsNull(RTotal(1)) Then
                            MskTotalEnvases.Text = RTotal(1)
                        Else
                            MskTotalEnvases.Text = "0"
                        End If
            Else
                MskTotalTarimas.Text = "0"
                MskTotalEnvases.Text = "0"
            End If
            
            
            If Err <> 0 Then
                MsgBox Err.Description
            End If
            
MousePointer = 0
        
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False

End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
            TxtBodega.Text = DBGridBusqueda.Columns(0)
            FrameBusqueda.Visible = False
            TxtBodega.SetFocus
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                    TxtBodega.Text = DBGridBusqueda.Columns(0)
                    FrameBusqueda.Visible = False
                    TxtBodega.SetFocus
            End If
End Sub

Private Sub DbGridFichaTecnica_HeadClick(ByVal ColIndex As Integer)
            RInventario.Sort = RInventario.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
                        
            CmdGenera_Click
End Sub


Private Sub Optbodega_Click()
            LblDescripcion.Visible = True
            TxtBodega.Visible = True
            TxtBodega.SetFocus
End Sub


Private Sub OptFichaBodega_Click()
            LblDescripcion.Visible = False
            LblBodega.Caption = ""
            TxtBodega.Visible = False
End Sub

Private Sub OptFichaTecnica_Click()
            LblDescripcion.Visible = False
            LblBodega.Caption = ""
            TxtBodega.Visible = False
End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select codigoBodega, Descripcion From BodegasInventario where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select codigoBodega, Descripcion From BodegasInventario where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select codigoBodega, Descripcion From BodegasInventario where CodigoBodega Like '%" & TxtBusqueda.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select codigoBodega, Descripcion From BodegasInventario where UPPER(CodigoBodega) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBodega_Change()
        If OptBodega.Value = True Then
            Set RBuscaBodega = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBodega.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBodega.Caption = RBuscaBodega!Descripcion
                Else
                    LblBodega.Caption = ""
                End If
        End If
            
End Sub

Private Sub TxtBodega_DblClick()
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
End Sub

Private Sub TxtBodega_GotFocus()
        TxtBodega.SelStart = 0
        TxtBodega.SelLength = Len(TxtBodega.Text)
End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
End Sub
