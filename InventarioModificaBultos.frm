VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioModificaBultos 
   Caption         =   "Modificacion De Bultos/Tarimas Inventario"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioModificaBultos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptOpc 
      Caption         =   "# Bulto/Tarima"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   45
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   11655
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   11
         Left            =   6120
         TabIndex        =   40
         Top             =   960
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   6120
         TabIndex        =   27
         Top             =   2040
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   9
         Left            =   6120
         TabIndex        =   26
         Top             =   1680
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   6120
         TabIndex        =   25
         Top             =   1320
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   6120
         TabIndex        =   24
         Top             =   600
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6120
         TabIndex        =   23
         Top             =   240
         Width           =   1100
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   22
         Top             =   2040
         Width           =   1400
      End
      Begin VB.TextBox TxtTexto 
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
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   21
         Top             =   1680
         Width           =   1400
      End
      Begin VB.TextBox TxtTexto 
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
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   20
         Top             =   1320
         Width           =   1400
      End
      Begin VB.TextBox TxtTexto 
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
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   19
         Top             =   960
         Width           =   1400
      End
      Begin VB.TextBox TxtTexto 
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
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   18
         Top             =   600
         Width           =   1400
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1400
      End
      Begin VB.Label lblFicTec 
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
         Height          =   615
         Left            =   2640
         TabIndex        =   44
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label LblBod 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7320
         TabIndex        =   43
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label LblOrdBol 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7320
         TabIndex        =   42
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Orden Boleta"
         Height          =   195
         Index           =   11
         Left            =   4800
         TabIndex        =   41
         Top             =   960
         Width           =   930
      End
      Begin VB.Label LblOrdPro 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7320
         TabIndex        =   39
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Index           =   10
         Left            =   4800
         TabIndex        =   38
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Entrada"
         Height          =   195
         Index           =   9
         Left            =   4800
         TabIndex        =   37
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Orden Produccion"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   36
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
         Height          =   195
         Index           =   7
         Left            =   4800
         TabIndex        =   35
         Top             =   600
         Width           =   555
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   34
         Top             =   240
         Width           =   525
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Batch"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Tarima"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Linea"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   390
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   450
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transaccion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Orden"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo De Saldo"
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton OptOpc 
         Caption         =   "> 0"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   " = 0"
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "< 0"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Todos"
         Height          =   195
         Index           =   5
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Txt2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Batch y Linea"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.OptionButton OptOpc 
      Caption         =   "Ficha Tecnica"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DbGrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            LCID            =   2058
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
            LCID            =   2058
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
      Caption         =   "&Salida"
      Height          =   975
      Left            =   10800
      Picture         =   "InventarioModificaBultos.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdVerDatos 
      Caption         =   "&Ver Datos"
      Height          =   975
      Left            =   9720
      Picture         =   "InventarioModificaBultos.frx":3D6C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Txt1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label LblFic 
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
      Left            =   5400
      TabIndex        =   9
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Lbl2 
      Caption         =   "Linea"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Lbl1 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1950
   End
End
Attribute VB_Name = "InventarioModificaBultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RVerDatos As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset
Dim RBuscaOrdenBoleta As New ADODB.Recordset
Dim RBuscaOrdenProduccion As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim VTexto As String


Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub CmdVerDatos_Click()
On Error Resume Next
        
MousePointer = 11
        Set RVerDatos = New ADODB.Recordset
                If OptOpc.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            VTexto = "Select D.* From DetalleEntradasInventario D Where D.FichaTecnica Like '" & UCase(Txt1.Text) & "%' And D.Linea Like '" & Txt2.Text & "%'"
                        Else 'ORACLE
                            VTexto = "Select D.* From DetalleEntradasInventario D Where UPPER(D.FichaTecnica) Like '" & UCase(Txt1.Text) & "%' And UPPER(D.Linea) Like '" & UCase(Txt2.Text) & "%'"
                        End If
                ElseIf OptOpc.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            VTexto = "Select D.* From DetalleEntradasInventario D Where D.Batch = " & Txt2.Text & " And D.Linea = '" & Txt1.Text & "'"
                        Else 'ORACLE
                            VTexto = "Select D.* From DetalleEntradasInventario D Where UPPER(D.Linea) = " & UCase(Txt1.Text) & " And UPPER(D.Batch) = '" & UCase(Txt2.Text) & "'"
                        End If
                ElseIf OptOpc.Item(6).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            VTexto = "Select D.* From DetalleEntradasInventario D Where D.OrdenProduccion Like '" & UCase(Txt1.Text) & "%'"
                        Else 'ORACLE
                            VTexto = "Select D.* From DetalleEntradasInventario D Where UPPER(D.OrdenProduccion) Like '" & UCase(Txt1.Text) & "%'"
                        End If
                ElseIf OptOpc.Item(7).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            VTexto = "Select D.* From DetalleEntradasInventario D Where D.Tarima = " & Txt1.Text
                        Else 'ORACLE
                            VTexto = "Select D.* From DetalleEntradasInventario D Where D.Tarima = " & Txt1.Text
                        End If
                
                End If
                
                If OptOpc.Item(2).Value = True Then
                            VTexto = VTexto & " And D.Saldo > 0 Order By D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima"
                ElseIf OptOpc.Item(3).Value = True Then
                            VTexto = VTexto & " And D.Saldo < 0 Order By D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima"
                ElseIf OptOpc.Item(4).Value = True Then
                            VTexto = VTexto & " And D.Saldo = 0 Order By D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima"
                ElseIf OptOpc.Item(5).Value = True Then
                            VTexto = VTexto & " Order By D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima"
                End If
                
                    Call Abrir_Recordset(RVerDatos, VTexto)
                    Set DbGrid.DataSource = RVerDatos
            
MousePointer = 0
                If Err <> 0 Then
                    MsgBox "ERROR " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If

End Sub

Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
                RVerDatos.Sort = RVerDatos.Fields(ColIndex).Name
End Sub


Private Sub DbGrid_SelChange(Cancel As Integer)
        TxtTexto.Item(0).Text = RVerDatos!Documento
        TxtTexto.Item(1).Text = RVerDatos!FechaProduccion
        TxtTexto.Item(2).Text = RVerDatos!Linea
        TxtTexto.Item(3).Text = RVerDatos!FichaTecnica
        TxtTexto.Item(4).Text = RVerDatos!Tarima
        TxtTexto.Item(5).Text = RVerDatos!Batch
        TxtTexto.Item(6).Text = RVerDatos!Calidad
        TxtTexto.Item(7).Text = RVerDatos!Bodega
        If IsNull(RVerDatos!OrdenProduccion) Then
            TxtTexto.Item(8).Text = ""
        Else
            TxtTexto.Item(8).Text = RVerDatos!OrdenProduccion
        End If
        TxtTexto.Item(9).Text = RVerDatos!CantidadEntrada
        TxtTexto.Item(10).Text = RVerDatos!Saldo
        If IsNull(RVerDatos!OrdenBoleta) Then
            TxtTexto.Item(11).Text = ""
        Else
            TxtTexto.Item(11).Text = RVerDatos!OrdenBoleta
        End If
        
        
        
        
End Sub

Private Sub OptOpc_Click(Index As Integer)
        If Index = 0 Then
            Lbl2.Caption = ""
            Txt2.Visible = False
            Lbl1.Caption = "Ficha Tecnica"
            Txt1.SetFocus
            Lbl2.Caption = "Linea"
            Txt2.Visible = True
        ElseIf Index = 1 Then
            Lbl2.Caption = "Batch"
            Lbl1.Caption = "Linea"
            Txt2.Visible = True
            Txt2.SetFocus
        ElseIf Index = 6 Then
            Lbl2.Caption = ""
            Txt2.Visible = False
            Lbl1.Caption = "Orden"
            Txt1.SetFocus
        ElseIf Index = 7 Then
            Lbl2.Caption = ""
            Txt2.Visible = False
            Lbl1.Caption = "# Bulto/Tarima"
            Txt1.SetFocus
        End If
End Sub

Private Sub Txt1_Change()
        If OptOpc.Item(0).Value = True Then
        
                 Set RBuscaFicha = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & Txt1.Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(Txt1.Text) & "'")
                     End If
                         If RBuscaFicha.RecordCount > 0 Then
                             LblFic.Caption = RBuscaFicha!Descrip
                         Else
                             LblFic.Caption = ""
                         End If
        ElseIf OptOpc.Item(6).Value = True Then
                     
                Set RBuscaOrden = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaOrden, "Select Documento From EncabezadoOrdenProduccion Where Documento = '" & Txt1.Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaOrden, "Select Documento From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(Txt1.Text) & "'")
                     End If
                         If RBuscaOrden.RecordCount > 0 Then
                             LblFic.Caption = RBuscaOrden!Documento
                         Else
                             LblFic.Caption = ""
                         End If
        End If
                     
            
        
End Sub

Private Sub Txt1_GotFocus()
        Txt1.SelStart = 0
        Txt1.SelLength = Len(Txt1.Text)
End Sub

Private Sub Txt1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Txt2_GotFocus()
        Txt2.SelStart = 0
        Txt2.SelLength = Len(Txt2.Text)
End Sub

Private Sub Txt2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        
        
    If Index = 7 Then
                Set RBuscaBodega = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtTexto.Item(7).Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtTexto.Item(7).Text) & "'")
                     End If
                         If RBuscaBodega.RecordCount > 0 Then
                             LblBod.Caption = RBuscaBodega!Descripcion
                         Else
                             LblBod.Caption = ""
                         End If
    ElseIf Index = 11 Then
                Set RBuscaOrdenBoleta = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaOrdenBoleta, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(11).Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaOrdenBoleta, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtTexto.Item(11).Text) & "'")
                     End If
                         If RBuscaOrdenBoleta.RecordCount > 0 Then
                             LblOrdBol.Caption = RBuscaOrdenBoleta!FichaTecnica
                         Else
                             LblOrdBol.Caption = ""
                         End If
    ElseIf Index = 8 Then
                Set RBuscaOrdenProduccion = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaOrdenProduccion, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(8).Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaOrdenProduccion, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtTexto.Item(8).Text) & "'")
                     End If
                         If RBuscaOrdenProduccion.RecordCount > 0 Then
                             LblOrdPro.Caption = RBuscaOrdenProduccion!FichaTecnica
                         Else
                             LblOrdPro.Caption = ""
                         End If
    ElseIf Index = 3 Then
                Set RBuscaFicha = New ADODB.Recordset
                     If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(3).Text & "'")
                     Else
                         Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(3).Text) & "'")
                     End If
                         If RBuscaFicha.RecordCount > 0 Then
                             lblFicTec.Caption = RBuscaFicha!Descrip
                         Else
                             lblFicTec.Caption = ""
                         End If

    End If
End Sub
