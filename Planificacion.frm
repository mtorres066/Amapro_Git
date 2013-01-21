VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Planificacion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Planificacion De Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Planificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8535
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   7335
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12938
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
         Height          =   735
         Left            =   10920
         Picture         =   "Planificacion.frx":2E7A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones de Reporte"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ficha Tecnica"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGrid 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13361
      _Version        =   393216
      Rows            =   5000
      Cols            =   5
      BackColor       =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11280
      Picture         =   "Planificacion.frx":4EEC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdAceptar 
      Height          =   495
      Left            =   10680
      Picture         =   "Planificacion.frx":6F5E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtFicTec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "signo '+' o doble click para ayuda"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LblFicTec 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Planificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBusMP As New ADODB.Recordset
Dim RInvPT As New ADODB.Recordset
Dim RInvMP As New ADODB.Recordset
Dim RPP As New ADODB.Recordset
Dim RPC As New ADODB.Recordset
Dim RPedPro As New ADODB.Recordset
Dim RPedCli As New ADODB.Recordset
Dim RSumaInvMP As New ADODB.Recordset
Dim RSumaInvPT As New ADODB.Recordset
Dim RSumaPedPro As New ADODB.Recordset
Dim RSumaPedCli As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RFichaTecnica As New ADODB.Recordset
Dim RPlanificacion As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim BFicha As Boolean
Dim BTipo As Boolean

Dim VFicha As String
Dim Cont As Integer

Private Sub CmdAceptar_Click()
On Error Resume Next
MousePointer = 11
        
        Cont = 1
        FGrid.Clear
        
        FGrid.Row = 0
        FGrid.Col = 0
        FGrid.ColWidth(0) = "100"
        FGrid.Col = 1
        FGrid.ColWidth(1) = "4800"
        FGrid.Col = 2
        FGrid.ColWidth(2) = "3300"
        FGrid.Col = 3
        FGrid.ColWidth(3) = "1400"
        FGrid.Col = 4
        FGrid.ColWidth(4) = "1700"
        
        'POR FICHA TECNICA
        If OptOpcion.Item(0).Value = True Then
            Set RFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RFichaTecnica, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '" & TxtFicTec.Text & "%' And Activa = -1")
                Else
                    Call Abrir_Recordset(RFichaTecnica, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '" & UCase(TxtFicTec.Text) & "%' And Activa = -1")
                End If
        'TIPO DE FICHA TECNICA
        ElseIf OptOpcion.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RFichaTecnica, "Select Esp_Tec, Descrip From FichaTecnica Where Tipo Like '" & TxtFicTec.Text & "%' And Activa = -1")
                Else
                    Call Abrir_Recordset(RFichaTecnica, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Tipo) Like '" & UCase(TxtFicTec.Text) & "%' And Activa = -1")
                End If
        End If
            
                If RFichaTecnica.RecordCount > 0 Then
                Else
                    MousePointer = 0
                    Exit Sub
                End If
                
                
                'CICLO DE TODAS LAS FICHA TECNICAS SELECCIONADAS
                Do Until RFichaTecnica.EOF
                
                                'CREA UNA LINEA DE COLORES PARA DIFERENCIAR LA OTRA FICHA TECNICA
                                                
                                Cont = Cont + 3
                                FGrid.Row = Cont
                                FGrid.Col = 1
                                FGrid.CellBackColor = &H80FF&
                                FGrid.CellFontBold = True
                                FGrid.Text = RFichaTecnica!Descrip
                                FGrid.Col = 2
                                FGrid.CellBackColor = &H80FF&
                                FGrid.CellFontBold = True
                                FGrid.CellFontSize = 12
                                FGrid.Text = RFichaTecnica!Esp_Tec
                                FGrid.Col = 3
                                FGrid.CellBackColor = &H80FF&
                                FGrid.Col = 4
                                FGrid.CellBackColor = &H80FF&
                                                                                
                                VFicha = RFichaTecnica!Esp_Tec
                
                                'MATERIAS PRIMAS TIENE ASIGNADA LA FICHA TECNICA
                                Set RBusMP = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusMP, "Select FT.Esp_Tec, FT.CodigoMateriaPrima, C.Descrip, C.UnidadMedida From FichaTecnicaConMateriaPrima FT, FichaTecnica C Where FT.Esp_Tec = '" & RFichaTecnica!Esp_Tec & "' And FT.CodigoMateriaPrima = C.Esp_Tec")
                                    Else
                                        Call Abrir_Recordset(RBusMP, "Select FT.Esp_Tec, FT.CodigoMateriaPrima, C.Descrip, C.UnidadMedida From FichaTecnicaConMateriaPrima FT, FichaTecnica C Where UPPER(FT.Esp_Tec) = '" & UCase(RFichaTecnica!Esp_Tec) & "' And UPPER(FT.CodigoMateriaPrima) = UPPER(C.Esp_Tec)")
                                    End If
                                    'If RBusMP.RecordCount > 0 Then
                                        
                                        
                                        'INVENTARIO PRODUCTO TERMINADO
                                        Set RInvPT = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RInvPT, "Select B.Descripcion, Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where DE.FichaTecnica = '" & VFicha & "' And DE.Saldo > 0 And DE.Bodega = B.CodigoBodega Group By DE.Bodega, B.Descripcion")
                                            Else
                                                Call Abrir_Recordset(RInvPT, "Select B.Descripcion, Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where UPPER(DE.FichaTecnica) = '" & UCase(VFicha) & "' And DE.Saldo > 0 And UPPER(DE.Bodega) = UPPER(B.CodigoBodega) Group By DE.Bodega, B.Descripcion")
                                            End If
                                        
                                        'SUMA EL INVENTARIO DE PRODUCTO TERMINADO
                                        Set RSumaInvPT = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RSumaInvPT, "Select Sum(Saldo) From DetalleEntradasInventario Where FichaTecnica = '" & VFicha & "' And Saldo > 0")
                                            Else
                                                Call Abrir_Recordset(RSumaInvPT, "Select Sum(Saldo) From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(VFicha) & "' And Saldo > 0")
                                            End If
                                                        Cont = Cont + 1
                                                        FGrid.Row = Cont
                                                        FGrid.Col = 1
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.CellFontBold = True
                                                        FGrid.CellFontSize = 12
                                                        FGrid.Text = "PRODUCTO TERMINADO"
                                                        FGrid.Col = 2
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.CellFontBold = True
                                                        'FGrid.Text = "INVENTARIO EN"
                                                        FGrid.Col = 3
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.CellFontBold = True
                                                        'FGrid.Text = "SALDO"
                                                        FGrid.Col = 4
                                                        FGrid.CellBackColor = vbCyan
                                            
                                            
                                            If RInvPT.RecordCount > 0 Then
                                                        
                                                'DESPLIEGA TODO EL INVENTARIO PRODUCTO TERMINADO
                                                Do Until RInvPT.EOF
                                                        Cont = Cont + 1
                                                        FGrid.Row = Cont
                                                        FGrid.Col = 2
                                                        FGrid.CellForeColor = &HFF0000
                                                        FGrid.Text = RInvPT(0)
                                                        FGrid.Col = 3
                                                        FGrid.CellForeColor = &HFF0000
                                                        FGrid.Text = Format(RInvPT(1), "#,###,##0")
                                                        FGrid.Col = 4
                                                        FGrid.CellForeColor = &HFF0000
                                                        FGrid.Text = "UNIDADES"
                                                        
                                                    RInvPT.MoveNext
                                                Loop
                                                
                                                    If RSumaInvPT.RecordCount > 0 Then
                                                        'DESPLIEGA EL TOTAL
                                                        Cont = Cont + 1
                                                        FGrid.Row = Cont
                                                        FGrid.Col = 2
                                                        FGrid.CellFontBold = True
                                                        FGrid.Text = "Total"
                                                        FGrid.Col = 3
                                                        FGrid.CellFontBold = True
                                                        FGrid.CellBackColor = &H8000000A
                                                        FGrid.Text = Format(RSumaInvPT(0), "#,###,##0")
                                                    End If
                                                        
                                            Else
                                                        Cont = Cont + 1
                                                        FGrid.Row = Cont
                                                        FGrid.Col = 1
                                                        FGrid.CellFontBold = True
                                                        FGrid.CellForeColor = vbRed
                                                        FGrid.Text = "NO HAY INVENTARIO"
                                                
                                            End If
                                            
                                                
                                                        'PEDIDOS DE PROVEEDORES ___________________________________________________________________________________________
                                                        Set RPedPro = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RPedPro, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Descripcion From DetallepedidosProveedores DP, EncabezadoPedidosProveedores EP, Proveedores P Where DP.Codigo = '" & VFicha & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Proveedor = P.CodigoProveedor Order By DP.FechaParaEntregar")
                                                            Else
                                                                Call Abrir_Recordset(RPedPro, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Descripcion From DetallepedidosProveedores DP, EncabezadoPedidosProveedores EP, Proveedores P Where UPPER(DP.Codigo) = '" & UCase(VFicha) & "' And DP.SaldoPorEntregar > 0 And UPPER(DP.Documento) = UPPER(EP.Documento) And UPPER(EP.Proveedor) = UPPER(P.CodigoProveedor) Order By DP.FechaParaEntregar")
                                                            End If
                                                        'SUMA TODOS LOS PEDIDOS DE PROVEEDORE
                                                        Set RSumaPedPro = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RSumaPedPro, "Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where Codigo = '" & VFicha & "' And SaldoPorEntregar > 0")
                                                            Else
                                                                Call Abrir_Recordset(RSumaPedPro, "Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where UPPER(Codigo) = '" & UCase(VFicha) & "' And SaldoPorEntregar > 0")
                                                            End If
                                                                
                                                            If RPedPro.RecordCount > 0 Then
                                                                Cont = Cont + 1
                                                                FGrid.Row = Cont
                                                                'FGrid.Col = 1
                                                                'FGrid.CellBackColor = vbYellow
                                                                'FGrid.Text = "PEDIDOS A PROVEEDORES"
                                                                FGrid.Col = 2
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "PEDIDOS A PROVEEDORES No."
                                                                FGrid.Col = 3
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "SALDO"
                                                                FGrid.Col = 4
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "FECHA ENTREGA"
                                                                
                                                            
                                                                Do Until RPedPro.EOF
                                                                        'DESPLIEGA EL DESGLOSE DE CADA PEDIDO
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 1
                                                                        FGrid.Text = RPedPro(3)
                                                                        FGrid.Col = 2
                                                                        FGrid.CellAlignment = 0
                                                                        FGrid.Text = RPedPro(0)
                                                                        FGrid.Col = 3
                                                                        FGrid.Text = Format(RPedPro(1), "#,###,##0.00")
                                                                        FGrid.Col = 4
                                                                        'CAMBIA EL COLOR SI LA FECHA ES MAYOR QUE LA RECEPCION
                                                                        If (Date > RPedPro(2)) Then
                                                                            FGrid.CellFontBold = True
                                                                            FGrid.CellForeColor = vbRed
                                                                        End If
                                                                        FGrid.Text = RPedPro(2)
                                                                    RPedPro.MoveNext
                                                                Loop
                                                                
                                                                If RSumaPedPro.RecordCount > 0 Then
                                                                        'DESPLIEGA EL TOTAL DE PEDIDOS
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 2
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.Text = "Total"
                                                                        FGrid.Col = 3
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.CellBackColor = &H8000000A
                                                                        FGrid.Text = Format(RSumaPedPro(0), "#,###,##0.00")
                                                                End If
                                                            Else
                                                                
                                                            End If
                                                        
                                                        'PEDIDOS DE CLIENTES ___________________________________________________________________________________________
                                                        Set RPedCli = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RPedCli, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes DP, EncabezadoPedidosClientes EP, Clientes C Where DP.Codigo = '" & VFicha & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Cliente = C.CodigoCliente Order By DP.FechaParaEntregar")
                                                            Else
                                                                Call Abrir_Recordset(RPedCli, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes DP, EncabezadoPedidosClientes EP, Clientes C Where UPPER(DP.Codigo) = '" & UCase(VFicha) & "' And DP.SaldoPorEntregar > 0 And UPPER(DP.Documento) = UPPER(EP.Documento) And UPPER(EP.Cliente) = UPPER(C.CodigoCliente) Order By DP.FechaParaEntregar")
                                                            End If
                                                            
                                                        'SUMA TODOS LOS PEDIDOS DE CLIENTES
                                                        Set RSumaPedCli = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RSumaPedCli, "Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where Codigo = '" & VFicha & "' And SaldoPorEntregar > 0")
                                                            Else
                                                                Call Abrir_Recordset(RSumaPedCli, "Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where UPPER(Codigo) = '" & UCase(VFicha) & "' And SaldoPorEntregar > 0")
                                                            End If
                                                        
                                                            If RPedCli.RecordCount > 0 Then
                                                                Cont = Cont + 1
                                                                FGrid.Row = Cont
                                                                'FGrid.Col = 1
                                                                'FGrid.CellBackColor = &H8080FF
                                                                'FGrid.CellFontBold = True
                                                                'FGrid.Text = "PEDIDOS DE CLIENTES"
                                                                FGrid.Col = 2
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "PEDIDOS DE CLIENTES No."
                                                                FGrid.Col = 3
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "SALDO"
                                                                FGrid.Col = 4
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "FECHA ENTREGA"
                                                           
                                                                'DESPLIEGA EL DESGLOSE DE CADA PEDIDO
                                                                Do Until RPedCli.EOF
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 1
                                                                        FGrid.Text = RPedCli(3)
                                                                        FGrid.Col = 2
                                                                        FGrid.CellAlignment = 0
                                                                        FGrid.Text = RPedCli(0)
                                                                        FGrid.Col = 3
                                                                        FGrid.Text = Format(RPedCli(1), "#,###,##0.00")
                                                                        FGrid.Col = 4
                                                                        'CAMBIA EL COLOR SI LA FECHA DE ENTREGA ES MAYOR
                                                                        If (Date > RPedCli(2)) Then
                                                                            FGrid.CellFontBold = True
                                                                            FGrid.CellForeColor = vbRed
                                                                        End If
                                                                        FGrid.Text = RPedCli(2)
                                                                    RPedCli.MoveNext
                                                                Loop
                                                                
                                                                If RSumaPedCli.RecordCount > 0 Then
                                                                        'DESPLIEGA EL TOTAL DE PEDIDOS
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 2
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.Text = "Total"
                                                                        FGrid.Col = 3
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.CellBackColor = &H8000000A
                                                                        FGrid.Text = Format(RSumaPedCli(0), "#,###,##0.00")
                                                                End If
                                                        
                                                            Else
                                                                
                                                            End If
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        '----------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        '-------------------------------------------------MATERIAS PRIMAS------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------
                        
                                                        Cont = Cont + 1
                                                        FGrid.Row = Cont
                                                        FGrid.Col = 1
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.CellFontBold = True
                                                        FGrid.CellFontSize = 12
                                                        FGrid.Text = "MATERIAS PRIMAS"
                                                        FGrid.Col = 2
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.Col = 3
                                                        FGrid.CellBackColor = vbCyan
                                                        FGrid.Col = 4
                                                        FGrid.CellBackColor = vbCyan
                                                        
                                                        
                                        'VERIFICA TODAS LAS MATERIAS PRIMAS
                                        Do Until RBusMP.EOF
                                                    'INVENTARIO MATERIA PRIMA
                                                    Set RInvMP = New ADODB.Recordset
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Call Abrir_Recordset(RInvMP, "Select B.Descripcion, Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where DE.FichaTecnica = '" & RBusMP(1) & "' And DE.Saldo > 0 And DE.Bodega = B.CodigoBodega Group By DE.Bodega, B.Descripcion")
                                                        Else
                                                            Call Abrir_Recordset(RInvMP, "Select B.Descripcion, Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where UPPER(DE.FichaTecnica) = '" & UCase(RBusMP(1)) & "' And DE.Saldo > 0 And UPPER(DE.Bodega) = UPPER(B.CodigoBodega) Group By DE.Bodega, B.Descripcion")
                                                        End If
                                                        
                                                    'SUMA EL TOTAL DEL INVENTARIO MATERIA PRIMA
                                                    Set RSumaInvMP = New ADODB.Recordset
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Call Abrir_Recordset(RSumaInvMP, "Select Sum(Saldo) From DetalleEntradasInventario Where Fichatecnica = '" & RBusMP(1) & "' and Saldo > 0")
                                                        Else
                                                            Call Abrir_Recordset(RSumaInvMP, "Select Sum(Saldo) From DetalleEntradasInventario Where UPPER(Fichatecnica) = '" & UCase(RBusMP(1)) & "' and Saldo > 0")
                                                        End If
                                                       
                                                            Cont = Cont + 1
                                                            FGrid.Row = Cont
                                                            FGrid.Col = 1
                                                            FGrid.CellBackColor = vbGreen
                                                            FGrid.CellFontBold = True
                                                            FGrid.Text = RBusMP(2)
                                                            FGrid.Col = 2
                                                            FGrid.CellBackColor = vbGreen
                                                            FGrid.CellFontBold = True
                                                         '   FGrid.Text = "INVENTARIO EN"
                                                            FGrid.Col = 3
                                                            FGrid.CellBackColor = vbGreen
                                                            FGrid.CellFontBold = True
                                                          '  FGrid.Text = "SALDO"
                                                            FGrid.Col = 4
                                                            FGrid.CellBackColor = vbGreen
                                                                
                                                            'DESPLIEGA TODOS LOS INVENTARIOS POR MATERIA PRIMA
                                                            Do Until RInvMP.EOF
                                                                    Cont = Cont + 1
                                                                    FGrid.Row = Cont
                                                                    FGrid.Col = 2
                                                                    FGrid.CellForeColor = &HFF0000
                                                                    FGrid.Text = RInvMP(0)
                                                                    FGrid.Col = 3
                                                                    FGrid.CellForeColor = &HFF0000
                                                                    FGrid.Text = Format(RInvMP(1), "#,###,##0.00")
                                                                    FGrid.Col = 4
                                                                    FGrid.CellForeColor = &HFF0000
                                                                    FGrid.Text = RBusMP(3)
                                                                RInvMP.MoveNext
                                                            Loop
                                                            
                                                            'DESPLIEGA TOTAL DE INVENTARIO DE MATERIA PRIMA
                                                            If RInvMP.RecordCount > 0 Then
                                                                    Cont = Cont + 1
                                                                    FGrid.Row = Cont
                                                                    FGrid.Col = 2
                                                                    'FGrid.CellBackColor = vbGreen
                                                                    FGrid.CellFontBold = True
                                                                    FGrid.Text = "Total"
                                                                    FGrid.Col = 3
                                                                    FGrid.CellFontBold = True
                                                                    FGrid.CellBackColor = &H8000000A
                                                                    FGrid.Text = Format(RSumaInvMP(0), "#,###,##0.00")
                                                            Else
                                                                Cont = Cont + 1
                                                                FGrid.Row = Cont
                                                                FGrid.Col = 1
                                                                FGrid.CellFontBold = True
                                                                FGrid.CellForeColor = vbRed
                                                                FGrid.Text = "NO HAY INVENTARIO"
                                                            End If
                                                        
                                                        'PEDIDOS DE PROVEEDORES ___________________________________________________________________________________________
                                                        Set RPedPro = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RPedPro, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Descripcion From DetallepedidosProveedores DP, EncabezadoPedidosProveedores EP, Proveedores P Where DP.Codigo = '" & RBusMP(1) & "' And DP.SaldoPorEntregar > 0 AND DP.Documento = EP.Documento And EP.Proveedor = P.CodigoProveedor Order By DP.FechaParaEntregar")
                                                            Else
                                                                Call Abrir_Recordset(RPedPro, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Descripcion From DetallepedidosProveedores DP, EncabezadoPedidosProveedores EP, Proveedores P Where UPPER(DP.Codigo) = '" & UCase(RBusMP(1)) & "' And DP.SaldoPorEntregar > 0 AND UPPER(DP.Documento) = UPPER(EP.Documento) And UPPER(EP.Proveedor) = UPPER(P.CodigoProveedor) Order By DP.FechaParaEntregar")
                                                            End If
                                                        'SUMA TODOS LOS PEDIDOS
                                                        Set RSumaPedPro = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RSumaPedPro, "Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where Codigo = '" & RBusMP(1) & "' And SaldoPorEntregar > 0")
                                                            Else
                                                                Call Abrir_Recordset(RSumaPedPro, "Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where UPPER(Codigo) = '" & UCase(RBusMP(1)) & "' And SaldoPorEntregar > 0")
                                                            End If
                                                        
                                                            If RPedPro.RecordCount > 0 Then
                                                                Cont = Cont + 1
                                                                FGrid.Row = Cont
                                                                'FGrid.Col = 1
                                                                'FGrid.CellFontBold = True
                                                                'FGrid.CellBackColor = vbYellow
                                                                'FGrid.Text = "PEDIDOS A PROVEEDORES"
                                                                FGrid.Col = 2
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "PEDIDOS A PROVEEDORES No."
                                                                FGrid.Col = 3
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "SALDO"
                                                                FGrid.Col = 4
                                                                FGrid.CellBackColor = vbYellow
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "FECHA ENTREGA"
                                                                
                                                                'DESPLIEGA EL DESGLOSE DE CADA PEDIDO
                                                                Do Until RPedPro.EOF
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 1
                                                                        FGrid.Text = RPedPro(3)
                                                                        FGrid.Col = 2
                                                                        FGrid.CellAlignment = 0
                                                                        FGrid.Text = RPedPro(0)
                                                                        FGrid.Col = 3
                                                                        FGrid.Text = Format(RPedPro(1), "#,###,##0.00")
                                                                        FGrid.Col = 4
                                                                        If (Date > RPedPro(2)) Then
                                                                            FGrid.CellFontBold = True
                                                                            FGrid.CellForeColor = vbRed
                                                                        End If
                                                                        FGrid.Text = RPedPro(2)
                                                                    RPedPro.MoveNext
                                                                Loop
                                                                        
                                                                If RSumaPedPro.RecordCount > 0 Then
                                                                        'DESPLIEGA EL TOTAL DE PEDIDOS
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 2
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.Text = "Total"
                                                                        FGrid.Col = 3
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.CellBackColor = &H8000000A
                                                                        FGrid.Text = Format(RSumaPedPro(0), "#,###,##0.00")
                                                                End If
                                                                
                                                            Else
                                                                
                                                            End If
                                                        
                                                        'PEDIDOS DE CLIENTES ___________________________________________________________________________________________
                                                        Set RPedCli = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RPedCli, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes DP, EncabezadoPedidosClientes EP, Clientes C Where DP.Codigo = '" & RBusMP(1) & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Cliente = C.CodigoCliente Order By DP.FechaParaEntregar")
                                                            Else
                                                                Call Abrir_Recordset(RPedCli, "Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes DP, EncabezadoPedidosClientes EP, Clientes C Where DP.Codigo = '" & RBusMP(1) & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Cliente = C.CodigoCliente Order By DP.FechaParaEntregar")
                                                            End If
                                                            
                                                        'SUMA PEDIDOS DE CLIENTES
                                                        Set RSumaPedCli = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RSumaPedCli, "Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where Codigo = '" & RBusMP(1) & "' And SaldoPorEntregar > 0")
                                                            Else
                                                                Call Abrir_Recordset(RSumaPedCli, "Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where UPPER(Codigo) = '" & UCase(RBusMP(1)) & "' And SaldoPorEntregar > 0")
                                                            End If
                                                        
                                                        
                                                            If RPedCli.RecordCount > 0 Then
                                                                Cont = Cont + 1
                                                                FGrid.Row = Cont
                                                                'FGrid.Col = 1
                                                                'FGrid.CellBackColor = &H8080FF
                                                                'FGrid.CellFontBold = True
                                                                'FGrid.Text = "PEDIDOS DE CLIENTES"
                                                                FGrid.Col = 2
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "PEDIDOS DE CLIENTES No."
                                                                FGrid.Col = 3
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "SALDO"
                                                                FGrid.Col = 4
                                                                FGrid.CellBackColor = &H8080FF
                                                                FGrid.CellFontBold = True
                                                                FGrid.Text = "FECHA ENTREGA"
                                                                
                                                                'DESPLIEGA EL DESGLOSE DE CADA PEDIDO
                                                                Do Until RPedCli.EOF
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 1
                                                                        FGrid.Text = RPedCli(3)
                                                                        FGrid.Col = 2
                                                                        FGrid.CellAlignment = 0
                                                                        FGrid.Text = RPedCli(0)
                                                                        FGrid.Col = 3
                                                                        FGrid.Text = Format(RPedCli(1), "#,###,##0.00")
                                                                        FGrid.Col = 4
                                                                        'CAMBIA EL COLOR SI LA FECHA DE ENTREGA ES MAYOR
                                                                        If (Date > RPedCli(2)) Then
                                                                            FGrid.CellFontBold = True
                                                                            FGrid.CellForeColor = vbRed
                                                                        End If
                                                                        FGrid.Text = RPedCli(2)
                                                                    RPedCli.MoveNext
                                                                Loop
                                                                
                                                                    If RSumaPedCli.RecordCount > 0 Then
                                                                        'DESPLIEGA EL TOTAL DE PEDIDOS
                                                                        Cont = Cont + 1
                                                                        FGrid.Row = Cont
                                                                        FGrid.Col = 2
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.Text = "Total"
                                                                        FGrid.Col = 3
                                                                        FGrid.CellFontBold = True
                                                                        FGrid.CellBackColor = &H8000000A
                                                                        FGrid.Text = Format(RSumaPedCli(0), "#,###,##0.00")
                                                                    End If
                                                        
                                                            Else
                                                                
                                                            End If
                                                        
                                                        
                                                RBusMP.MoveNext
                                                
                                                             
                                        Loop
                                        
                                        
                    RFichaTecnica.MoveNext 'CICLO DE TODAS LAS FICHAS TECNICAS
                Loop
                                    
                            
            FGrid.SetFocus
MousePointer = 0
End Sub

Private Sub CmdSale_Click()
        FrameConsultas.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        TxtFicTec.Text = DbGridBusqueda.Columns(0)
        TxtFicTec.SetFocus
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        TxtFicTec.Text = DbGridBusqueda.Columns(0)
        TxtFicTec.SetFocus
        FrameConsultas.Visible = False
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            Lbl.Caption = "Ficha Tecnica"
        ElseIf Index = 1 Then
            Lbl.Caption = "Tipo Ficha Tecnica"
        End If
End Sub

Private Sub TxtConsultas_Change()
    Set RBusqueda = New ADODB.Recordset

    If BFicha = True Then
            'FICHA TECNICA
            If OptDes.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Descrip Like '%" & TxtConsultas.Text & "%' And Activa = -1")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtConsultas.Text) & "%' And Activa = -1")
                End If
            ElseIf OptCod.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Esp_Tec Like '%" & TxtConsultas.Text & "%' And Activa = -1")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtConsultas.Text) & "%' And Activa = -1")
                End If
            End If
     Else
        'FICHA TECNICA TIPOS
            If OptDes.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where Descripcion Like '%" & TxtConsultas.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where UPPER(Descripcion) Like '%" & UCase(TxtConsultas.Text) & "%'")
                End If
            ElseIf OptCod.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where CodigoTipo Like '%" & TxtConsultas.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) Like '%" & UCase(TxtConsultas.Text) & "%'")
                End If
                
            End If
     End If
        
        
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "5000"
    
End Sub

Private Sub TxtConsultas_GotFocus()
        TxtConsultas.SelStart = 0
        TxtConsultas.SelLength = Len(TxtConsultas.Text)
End Sub

Private Sub TxtConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtFicTec_Change()
    If OptOpcion.Item(0).Value = True Then
        Set RBuscaFicha = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            
            If RBuscaFicha.RecordCount > 0 Then
                LblFicTec.Caption = RBuscaFicha!Descrip
            Else
                LblFicTec.Caption = ""
            End If
    ElseIf OptOpcion.Item(1).Value = True Then
        Set RBuscaFicha = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtFicTec.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RBuscaFicha.RecordCount > 0 Then
                LblFicTec.Caption = RBuscaFicha!Descripcion
            Else
                LblFicTec.Caption = ""
            End If
    End If
            
End Sub

Private Sub TxtFicTec_DblClick()
        Set RBusqueda = New ADODB.Recordset
        If OptOpcion.Item(0).Value = True Then
            BFicha = True
            BTipo = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
        ElseIf OptOpcion.Item(1).Value = True Then
            BFicha = False
            BTipo = True
            Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
        End If
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "5000"
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
                If OptOpcion.Item(0).Value = True Then
                    BFicha = True
                    BTipo = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                ElseIf OptOpcion.Item(1).Value = True Then
                    BFicha = False
                    BTipo = True
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
                End If
                    
                    Set DbGridBusqueda.DataSource = RBusqueda
                    DbGridBusqueda.Columns(1).Width = "5000"
                    FrameConsultas.Visible = True
                    TxtConsultas.SetFocus
    End If
    
    

End Sub
