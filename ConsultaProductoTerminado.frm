VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ConsultaProductoTerminado 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Planificacion De Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ConsultaProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Reporte"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Tipo Ficha Tecnica"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   10200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8535
      Left            =   6960
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   10920
         Picture         =   "ConsultaProductoTerminado.frx":2E7A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   7320
         TabIndex        =   7
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Data DataConsultas 
         Caption         =   "consultas"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "ConsultaProductoTerminado.frx":4EEC
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "ConsultaProductoTerminado.frx":4F08
         TabIndex        =   13
         ToolTipText     =   "Signo '+' o Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   11535
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
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
      Rows            =   200
      Cols            =   5
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11280
      Picture         =   "ConsultaProductoTerminado.frx":58E3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdAceptar 
      Height          =   495
      Left            =   10680
      Picture         =   "ConsultaProductoTerminado.frx":7955
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
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "signo '+' o doble click para ayuda"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LblFicTec 
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
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Width           =   1230
   End
End
Attribute VB_Name = "ConsultaProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBusMP As Recordset
Dim RInvPT As Recordset
Dim RInvMP As Recordset
Dim RPP As Recordset
Dim RPC As Recordset
Dim RPedPro As Recordset
Dim RPedCli As Recordset

Dim RSumaInvMP As Recordset
Dim RSumaInvPT As Recordset
Dim RSumaPedPro As Recordset
Dim RSumaPedCli As Recordset

Dim RBuscaFicha As Recordset


Dim VFicha As String
Dim cont As Integer

Private Sub CmdAceptar_Click()
On Error Resume Next
VFicha = TxtFicTec.Text
        
        
        
        cont = 1
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
        

        
        'MATERIAS PRIMAS TIENE ASIGNADA LA FICHA TECNICA
        Set RBusMP = Db.OpenRecordset("Select FT.Esp_Tec, FT.CodigoMateriaPrima, C.Descripcion, C.UnidadMedida From FichaTecnicaConMateriaPrima AS FT, CorrelativosMateriaPrima as C Where FT.Esp_Tec = '" & TxtFicTec.Text & "' And FT.CodigoMateriaPrima = C.CodigoMateriaPrima")
            'If RBusMP.RecordCount > 0 Then
                
                
                'INVENTARIO PRODUCTO TERMINADO
                Set RInvPT = Db.OpenRecordset("Select B.Descripcion, Sum(DE.Saldo) From DetalleEntradasProductoTerminado as DE, BodegasProductoTerminado as B Where DE.FichaTecnica = '" & VFicha & "' And DE.Saldo > 0 And DE.Bodega = B.CodigoBodega Group By DE.Bodega, B.Descripcion")
                'SUMA EL INVENTARIO DE PRODUCTO TERMINADO
                Set RSumaInvPT = Db.OpenRecordset("Select Sum(Saldo) From DetalleEntradasProductoTerminado Where FichaTecnica = '" & VFicha & "' And Saldo > 0")
                  
                                FGrid.Row = cont
                                FGrid.Col = 1
                                FGrid.CellBackColor = vbCyan
                                FGrid.CellFontBold = True
                                FGrid.CellFontSize = 12
                                FGrid.Text = "PRODUCTO TERMINADO"
                                FGrid.Col = 2
                                FGrid.CellBackColor = vbCyan
                                FGrid.CellFontBold = True
                                FGrid.Text = "INVENTARIO EN"
                                FGrid.Col = 3
                                FGrid.CellBackColor = vbCyan
                                FGrid.CellFontBold = True
                                FGrid.Text = "SALDO"
                                FGrid.Col = 4
                                FGrid.CellBackColor = vbCyan
                    
                    
                    If RInvPT.RecordCount > 0 Then
                                
                        'DESPLIEGA TODO EL INVENTARIO PRODUCTO TERMINADO
                        Do Until RInvPT.EOF
                                cont = cont + 1
                                FGrid.Row = cont
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
                                cont = cont + 1
                                FGrid.Row = cont
                                FGrid.Col = 2
                                FGrid.CellFontBold = True
                                FGrid.Text = "Total"
                                FGrid.Col = 3
                                FGrid.CellFontBold = True
                                FGrid.CellBackColor = &H8000000A
                                FGrid.Text = Format(RSumaInvPT(0), "#,###,##0")
                            End If
                                
                    Else
                                cont = cont + 1
                                FGrid.Row = cont
                                FGrid.Col = 1
                                FGrid.CellFontBold = True
                                FGrid.CellForeColor = vbRed
                                FGrid.Text = "NO HAY INVENTARIO"
                        
                    End If
                    
                        
'PEDIDOS DE PROVEEDORES ___________________________________________________________________________________________
                                Set RPedPro = Db.OpenRecordset("Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Proveedor From DetallepedidosProveedores As DP, EncabezadoPedidosProveedores as EP, Proveedores as P Where DP.Codigo = '" & VFicha & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Proveedor = P.CodigoProveedor Order By DP.FechaParaEntregar")
                                'SUMA TODOS LOS PEDIDOS DE PROVEEDORE
                                Set RSumaPedPro = Db.OpenRecordset("Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where Codigo = '" & VFicha & "' And SaldoPorEntregar > 0")
                                    If RPedPro.RecordCount > 0 Then
                                        cont = cont + 1
                                        FGrid.Row = cont
                                        FGrid.Col = 1
                                        FGrid.CellBackColor = vbYellow
                                        FGrid.Text = "PEDIDOS A PROVEEDORES"
                                        FGrid.Col = 2
                                        FGrid.CellBackColor = vbYellow
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDO No."
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                Set RPedCli = Db.OpenRecordset("Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes AS DP, EncabezadoPedidosClientes as EP, Clientes as C Where DP.Codigo = '" & VFicha & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Cliente = C.CodigoCliente Order By DP.FechaParaEntregar")
                                'SUMA TODOS LOS PEDIDOS DE CLIENTES
                                Set RSumaPedCli = Db.OpenRecordset("Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where Codigo = '" & VFicha & "' And SaldoPorEntregar > 0")
                                
                                    If RPedCli.RecordCount > 0 Then
                                        cont = cont + 1
                                        FGrid.Row = cont
                                        FGrid.Col = 1
                                        FGrid.CellBackColor = &H8080FF
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDOS DE CLIENTES"
                                        FGrid.Col = 2
                                        FGrid.CellBackColor = &H8080FF
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDO No."
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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

                                cont = cont + 1
                                FGrid.Row = cont
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
                            Set RInvMP = Db.OpenRecordset("Select B.Descripcion, Sum(DE.SaldoDisponibilidad) From DetalleEntradasMateriaPrima as DE, BodegasMateriaPrima as B Where DE.Codigo = '" & RBusMP(1) & "' And DE.SaldoDisponibilidad > 0 And DE.BodegaDisponibilidad = B.CodigoBodega Group By DE.BodegaDisponibilidad, B.Descripcion")
                            'SUMA EL TOTAL DEL INVENTARIO MATERIA PRIMA
                            Set RSumaInvMP = Db.OpenRecordset("Select Sum(SaldoDisponibilidad) From DetalleEntradasMateriaPrima Where Codigo = '" & RBusMP(1) & "' and SaldoDisponibilidad > 0")
                               
                                    cont = cont + 1
                                    FGrid.Row = cont
                                    FGrid.Col = 1
                                    FGrid.CellBackColor = vbGreen
                                    FGrid.CellFontBold = True
                                    FGrid.Text = RBusMP(2)
                                    FGrid.Col = 2
                                    FGrid.CellBackColor = vbGreen
                                    FGrid.CellFontBold = True
                                    FGrid.Text = "INVENTARIO EN"
                                    FGrid.Col = 3
                                    FGrid.CellBackColor = vbGreen
                                    FGrid.CellFontBold = True
                                    FGrid.Text = "SALDO"
                                    FGrid.Col = 4
                                    FGrid.CellBackColor = vbGreen
                                        
                                    'DESPLIEGA TODOS LOS INVENTARIOS POR MATERIA PRIMA
                                    Do Until RInvMP.EOF
                                            cont = cont + 1
                                            FGrid.Row = cont
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
                                            cont = cont + 1
                                            FGrid.Row = cont
                                            FGrid.Col = 2
                                            'FGrid.CellBackColor = vbGreen
                                            FGrid.CellFontBold = True
                                            FGrid.Text = "Total"
                                            FGrid.Col = 3
                                            FGrid.CellFontBold = True
                                            FGrid.CellBackColor = &H8000000A
                                            FGrid.Text = Format(RSumaInvMP(0), "#,###,##0.00")
                                    Else
                                        cont = cont + 1
                                        FGrid.Row = cont
                                        FGrid.Col = 1
                                        FGrid.CellFontBold = True
                                        FGrid.CellForeColor = vbRed
                                        FGrid.Text = "NO HAY INVENTARIO"
                                    End If
                                
'PEDIDOS DE PROVEEDORES ___________________________________________________________________________________________
                                Set RPedPro = Db.OpenRecordset("Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, P.Proveedor From DetallepedidosProveedores as DP, EncabezadoPedidosProveedores as EP, Proveedores as P Where DP.Codigo = '" & RBusMP(1) & "' And DP.SaldoPorEntregar > 0 AND DP.Documento = EP.Documento And EP.Proveedor = P.CodigoProveedor Order By DP.FechaParaEntregar")
                                'SUMA TODOS LOS PEDIDOS
                                Set RSumaPedPro = Db.OpenRecordset("Select SUM(SaldoPorEntregar) From DetallepedidosProveedores Where Codigo = '" & RBusMP(1) & "' And SaldoPorEntregar > 0")
                                
                                    If RPedPro.RecordCount > 0 Then
                                        cont = cont + 1
                                        FGrid.Row = cont
                                        FGrid.Col = 1
                                        FGrid.CellFontBold = True
                                        FGrid.CellBackColor = vbYellow
                                        FGrid.Text = "PEDIDOS A PROVEEDORES"
                                        FGrid.Col = 2
                                        FGrid.CellBackColor = vbYellow
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDO No."
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                Set RPedCli = Db.OpenRecordset("Select DP.Documento, DP.SaldoPorEntregar, DP.FechaParaEntregar, C.Descripcion From DetallepedidosClientes as DP, EncabezadoPedidosClientes as EP, Clientes as C Where DP.Codigo = '" & RBusMP(1) & "' And DP.SaldoPorEntregar > 0 And DP.Documento = EP.Documento And EP.Cliente = C.CodigoCliente Order By DP.FechaParaEntregar")
                                'SUMA PEDIDOS DE CLIENTES
                                Set RSumaPedCli = Db.OpenRecordset("Select SUM(SaldoPorEntregar) From DetallepedidosClientes Where Codigo = '" & RBusMP(1) & "' And SaldoPorEntregar > 0")
                                
                                
                                    If RPedCli.RecordCount > 0 Then
                                        cont = cont + 1
                                        FGrid.Row = cont
                                        FGrid.Col = 1
                                        FGrid.CellBackColor = &H8080FF
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDOS DE CLIENTES"
                                        FGrid.Col = 2
                                        FGrid.CellBackColor = &H8080FF
                                        FGrid.CellFontBold = True
                                        FGrid.Text = "PEDIDO No."
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
                                                cont = cont + 1
                                                FGrid.Row = cont
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
            'Else
            '    MsgBox "La Ficha Tecnica No Tiene Asignadas Materias Primas", vbOKOnly + vbInformation, "Informacion"
            'End If
    
            FGrid.SetFocus
End Sub


Private Sub CmdSale_Click()
        FrameConsultas.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridConsultas_DblClick()
        TxtFicTec.Text = DBGridConsultas.Columns(0)
        TxtFicTec.SetFocus
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        TxtFicTec.Text = DBGridConsultas.Columns(0)
        TxtFicTec.SetFocus
        FrameConsultas.Visible = False
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
    DataConsultas.Connect = GConnect
    DataConsultas.DatabaseName = BasedeDatos

End Sub

Private Sub TxtConsultas_Change()
        'FICHA TECNICA
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Descrip Like '" & TxtConsultas.Text & "*' And Activa = -1"
            Else
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Descrip Like '*" & TxtConsultas.Text & "*' And Activa = -1"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Esp_Tec Like '" & TxtConsultas.Text & "*' And Activa = -1"
            Else
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Esp_Tec Like '*" & TxtConsultas.Text & "*' And Activa = -1"
            End If
        End If
    
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "5000"
    
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
        Set RBuscaFicha = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            If RBuscaFicha.RecordCount > 0 Then
                LblFicTec.Caption = RBuscaFicha!Descrip
            Else
                LblFicTec.Caption = ""
            End If
            
End Sub

Private Sub TxtFicTec_DblClick()
        DataConsultas.RecordSource = "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Activa = -1"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        DBGridConsultas.Columns(1).Width = "5000"
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
        DataConsultas.RecordSource = "Select Esp_Tec, Descrip, Size, Envases From FichaTecnica"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        DBGridConsultas.Columns(1).Width = "5000"
        FrameConsultas.Visible = True
        TxtConsultas.SetFocus
    End If
    
    

End Sub
