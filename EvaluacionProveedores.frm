VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EvaluacionProveedores 
   Caption         =   "Evaluacion De Proveedores"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "EvaluacionProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8910
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
      Height          =   5652
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4455
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         Left            =   8160
         Picture         =   "EvaluacionProveedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtPro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo De Reporte"
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
      Height          =   735
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   3375
      Begin VB.OptionButton OptRep 
         Caption         =   "Certificado"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptRep 
         Caption         =   "Detalle"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptRep 
         Caption         =   "Resumen"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   855
      Left            =   7320
      Picture         =   "EvaluacionProveedores.frx":293C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1500
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Height          =   855
      Left            =   5760
      Picture         =   "EvaluacionProveedores.frx":49AE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones De Reporte"
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
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Grupo Proveedor"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox MskFecEva 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DTPFecFin 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20643843
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20643843
      CurrentDate     =   37798
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde "
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
      Left            =   1080
      TabIndex        =   23
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fechas De Entrega"
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
      Index           =   4
      Left            =   1080
      TabIndex        =   21
      Top             =   1320
      Width           =   1650
   End
   Begin VB.Label LblPro 
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Lbllabel 
      AutoSize        =   -1  'True
      Caption         =   "Codigo Proveedor"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha A Evaluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1425
   End
End
Attribute VB_Name = "EvaluacionProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VFechaEvaluar As Date
Dim VFechaInicial As Date
Dim VFechaFinal As Date
Dim VFechaParaEntregar As Date
Dim VFechaEntregadoTotal As String

Dim VDiasAtrasoEntrega As Integer

Dim VPorcentajeTiempo As Integer
Dim VPorcentajeTiempo2 As Integer
Dim VPorcentajeCantidad As Integer
Dim VPorcentajeCantidad2 As Integer
Dim VPorcentajeCalidad As Integer
Dim VPorcentajeCalidad2 As Integer
Dim VTotal As Integer

Dim VEntregado As Single
Dim VPedido As Single
Dim VSaldo As Single

Dim RProveedores As New ADODB.Recordset
Dim RBuscaProveedor As New ADODB.Recordset
Dim RPedidos As New ADODB.Recordset
Dim RPedidosNoConforme As New ADODB.Recordset
Dim REvaluacionProveedores As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim VTexto As String

Dim BProveedor As Boolean
Dim BGrupo As Boolean


Private Sub CmdImprimir_Click()
On Error Resume Next

MousePointer = 11
        'SELECCIONAMOS LOS PROVEEDORES A BUSCAR
        'POR PROVEEDOR
        If OptOpcion.Item(0).Value = True Then
                Set RProveedores = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor From Proveedores Where CodigoProveedor Like '" & TxtPro.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor From Proveedores Where UPPER(CodigoProveedor) Like '" & UCase(TxtPro.Text) & "%'")
                    End If
                    If RProveedores.RecordCount > 0 Then
                    
                    End If
        'POR GRUPO DE PROVEEDORES
        Else
                Set RProveedores = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProveedores, "Select * From Proveedores Where Grupo = '" & TxtPro.Text & "'")
                    Else
                        Call Abrir_Recordset(RProveedores, "Select * From Proveedores Where UPPER(Grupo) = '" & UCase(TxtPro.Text) & "'")
                    End If
                    If RProveedores.RecordCount > 0 Then
                    
                    End If
        End If
        
        
        'BORRA TODOS LOS DATOS DE LA BASE DE DATOS PARA SACAR UN NUEVO REPORTE
        Conexion.Execute ("Delete From ReporteEvaluacionProveedores")
        
        'GUARDAMOS LAS VARIABLES DE FECHA
        VFechaEvaluar = MskFecEva.Text
        VFechaInicial = DtpFecIni.Value
        VFechaFinal = DTPFecFin.Value
        
                    'CREAMOS UN CICLO CON TODOS LOS CODIGOS DE PROVEEDORES DEPENDIENDO DE LA OPCION MARCADA
                    Do Until RProveedores.EOF
                                'BUSCAMOS LOS PEDIDOS QUE TENGA UN PROVEEDOR PERO EN EL RANGO DE FECHAS DE ENTREGA
                                Set RPedidos = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RPedidos, "Select DP.* From DetallePedidosProveedores DP, EncabezadoPedidosProveedores EP Where EP.Proveedor = '" & RProveedores!CodigoProveedor & "' AND EP.Documento = DP.Documento And DP.FechaParaEntregar >= #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And DP.FechaParaEntregar <= #" & Format(VFechaFinal, "mm/dd/yyyy") & "# Order by DP.FechaParaEntregar, DP.Documento")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RPedidos, "Select DP.* From DetallePedidosProveedores DP, EncabezadoPedidosProveedores EP Where UPPER(EP.Proveedor) = '" & UCase(RProveedores!CodigoProveedor) & "' AND UPPER(EP.Documento) = UPPER(DP.Documento) And DP.FechaParaEntregar >= To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And DP.FechaParaEntregar <= To_Date('" & VFechaFinal & "', 'dd/mm/yyyy')" & " Order by DP.FechaParaEntregar, DP.Documento")
                                    End If
                                    If RPedidos.RecordCount > 0 Then
                                        'CREA UN CICLO CON TODOS LOS PEDIDOS
                                        Do Until RPedidos.EOF
                                        
                                                                                                
                                                VFechaParaEntregar = RPedidos!FechaParaEntregar
                                                VEntregado = RPedidos!CantidadEntregada
                                                VPedido = RPedidos!CantidadPedido
                                                VSaldo = RPedidos!SaldoPorEntregar
                                                
                                                                                                
                                                'SI EL SALDO DEL PEDIDO ES IGUAL O MENOR QUE CERO ENTONCES EL PEDIDO
                                                'YA SE ENTREGO Y SI HAY UNA FECHA DE ENTREGA TOTAL
                                                'DE LO CONTRARIO SI TIENE SALDO ES PORQUE NO HAY FECHA DE ENTREGA TOTAL
                                                'ENTONCES UTILIZAMOS COMO ENTREGA TOTAL LA FECHA A EVALUAR
                                                'If VSaldo <= 0 Then
                                                        If IsNull(RPedidos!FechaEntregaTotal) Then
                                                            VFechaEntregadoTotal = ""
                                                        Else
                                                            VFechaEntregadoTotal = RPedidos!FechaEntregaTotal
                                                        End If
                                                'Else
                                                '        VFechaEntregadoTotal = ""
                                                'End If
                                                
                                                
                                                '% DE TIEMPO
                                                                'If VFechaEvaluar < VFechaParaEntregar Then
                                                                    'NO HACE NADA
                                                                '        VPorcentajeTiempo = 0
                                                                'If VSaldo <= 0 Then
                                                                        If VFechaEntregadoTotal = "" Then
                                                                            VDiasAtrasoEntrega = (DateValue(VFechaEvaluar) - DateValue(VFechaParaEntregar))
                                                                        Else
                                                                            VDiasAtrasoEntrega = (DateValue(VFechaEntregadoTotal) - DateValue(VFechaParaEntregar))
                                                                        End If
                                                                'Else
                                                                '        VDiasAtrasoEntrega = (DateValue(VFechaEvaluar) - DateValue(VFechaParaEntregar))
                                                                'End If
                                                                
                                                                        'SI LA VARIABLE VDIASDEATRASO ES MENOR QUE CERO ES PORQUE ENTREGO EL PEDIDO ANTES DE LA FECHA
                                                                        If VDiasAtrasoEntrega < 0 Then
                                                                            VDiasAtrasoEntrega = 0
                                                                        End If
                                
                                                                
                                                                        'CALCULA EL PORCENTAJE DEPENDIENDO DEL RANGO DE DIAS
                                                                        If VDiasAtrasoEntrega = 0 Then
                                                                            VPorcentajeTiempo = 100
                                                                        ElseIf VDiasAtrasoEntrega <= 5 Then
                                                                            VPorcentajeTiempo = 95
                                                                        ElseIf VDiasAtrasoEntrega >= 6 And VDiasAtrasoEntrega <= 10 Then
                                                                            VPorcentajeTiempo = 90
                                                                        ElseIf VDiasAtrasoEntrega >= 11 And VDiasAtrasoEntrega <= 15 Then
                                                                            VPorcentajeTiempo = 85
                                                                        ElseIf VDiasAtrasoEntrega >= 16 And VDiasAtrasoEntrega <= 20 Then
                                                                            VPorcentajeTiempo = 80
                                                                        ElseIf VDiasAtrasoEntrega >= 21 And VDiasAtrasoEntrega <= 25 Then
                                                                            VPorcentajeTiempo = 75
                                                                        ElseIf VDiasAtrasoEntrega >= 26 And VDiasAtrasoEntrega <= 30 Then
                                                                            VPorcentajeTiempo = 70
                                                                        ElseIf VDiasAtrasoEntrega > 31 Then
                                                                            VPorcentajeTiempo = 0
                                                                        End If
                                                                
                                              '% DE CANTIDAD
                                                                If VFechaEvaluar < VFechaParaEntregar Then
                                                                    'NO HACE NADA
                                                                        VPorcentajeCantidad = 0
                                                                Else
                                                                        VPorcentajeCantidad = ((VEntregado / VPedido) * 100)
                                                                        
                                                                        'SI YA LLEVA UN 90 % DE ENTREGA O MAS ES COMO QUE YA HUBIERA ENTREGADO TODO
                                                                        If VPorcentajeCantidad >= 90 Then
                                                                            VPorcentajeCantidad = 100
                                                                        End If
                                                                        
                                                                End If
                                             
                                             '% DE CALIDAD
                                                                'BUSCAMOS EL PROMEDIO DEL PORCENTAJE DE NO CONFORMIDAD POR PEDIDO Y CODIGO
                                                                Set RPedidosNoConforme = New ADODB.Recordset
                                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                                        Call Abrir_Recordset(RPedidosNoConforme, "Select avg(PorcentajeConforme) From PedidosProveedoresPorcentajeNo Where Pedido = '" & RPedidos!Documento & "' And Codigo = '" & RPedidos!Codigo & "'")
                                                                    Else 'ORACLE
                                                                        Call Abrir_Recordset(RPedidosNoConforme, "Select avg(PorcentajeConforme) From PedidosProveedoresPorcentajeNo Where UPPER(Pedido) = '" & UCase(RPedidos!Documento) & "' And UPPER(Codigo) = '" & UCase(RPedidos!Codigo) & "'")
                                                                    End If
                                                                    If RPedidosNoConforme.RecordCount > 0 Then
                                                                        If IsNull(RPedidosNoConforme(0)) Then
                                                                            VPorcentajeCalidad = 0
                                                                        Else
                                                                            VPorcentajeCalidad = RPedidosNoConforme(0)
                                                                        End If
                                                                    Else
                                                                        VPorcentajeCalidad = 0
                                                                    End If
                                                                    
                                                                    'ASIGNA VALORES A VARIABLES EN BASE A PROMEDIO PONDERADO PARA SACAR EL TOTAL
                                                                    'ASIGNANDOLES UN FACTOR POR PRIORIDAD A LOS PORCENTAJES
                                                                    VPorcentajeTiempo2 = VPorcentajeTiempo * 0.3
                                                                    VPorcentajeCantidad2 = VPorcentajeCantidad * 0.3
                                                                    VPorcentajeCalidad2 = VPorcentajeCalidad * 0.4
                                                                    VTotal = VPorcentajeTiempo2 + VPorcentajeCantidad2 + VPorcentajeCalidad2
                                                                    
                                                                'AGREGAMOS DATOS A LA BASE DE DATOS
                                                                    VTexto = "'" & RProveedores!CodigoProveedor & "', '"
                                                                    VTexto = VTexto & RPedidos!Documento & "', '"
                                                                    VTexto = VTexto & RPedidos!Codigo & "', "
                                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                                         VTexto = VTexto & "#" & Format(VFechaParaEntregar, "mm/dd/yyyy") & "#, '" 'FECHA
                                                                    Else 'ORACLE
                                                                         VTexto = VTexto & "To_Date('" & VFechaParaEntregar & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                                                                    End If
                                                                    VTexto = VTexto & Format(VFechaEntregadoTotal, "dd/mm/yyyy") & "', "
                                                                    VTexto = VTexto & VDiasAtrasoEntrega & ", "
                                                                    VTexto = VTexto & VPedido & ", "
                                                                    VTexto = VTexto & VEntregado & ", "
                                                                    VTexto = VTexto & VSaldo & ", "
                                                                    VTexto = VTexto & VPorcentajeTiempo & ", "
                                                                    VTexto = VTexto & VPorcentajeCantidad & ", "
                                                                    VTexto = VTexto & VPorcentajeCalidad & ", "
                                                                    VTexto = VTexto & VTotal
                                                                    
                                                                    Conexion.Execute "insert Into ReporteEvaluacionProveedores Values(" & VTexto & ")"
                                                                       
                                                                If Err <> 0 Then
                                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                                                End If
                                                                                                            
                                            RPedidos.MoveNext
                                        Loop
                                    End If
                        RProveedores.MoveNext
                    Loop
                    
                    If OptOpcion.Item(0).Value = True Then
                        GTituloReporte = "Desde Fecha De Entrega " & Format(DtpFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPFecFin.Value, "dd/mm/yyyy") & " Y Fecha Evaluacion " & MskFecEva.Text & " Por Proveedor " & TxtPro.Text & " " & LblPro.Caption & "'"
                    Else
                        GTituloReporte = "Desde Fecha De Entrega " & Format(DtpFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPFecFin.Value, "dd/mm/yyyy") & " Y Fecha Evaluacion " & MskFecEva.Text & " Por Grupo " & TxtPro.Text & " " & LblPro.Caption & "'"
                    End If
                        GCriteriaReporte = ""
                    
                    'MUESTRA EL REPORTE
                    If GOrigenDeDatos = "AmaproAccess" Then
                        If OptRep.Item(0).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresResumen.rpt"
                        ElseIf OptRep.Item(1).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresDetalle.rpt"
                        ElseIf OptRep.Item(2).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresCertificado.rpt"
                        End If
                    Else 'ORACLE
                        If OptRep.Item(0).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresResumenO.rpt"
                        ElseIf OptRep.Item(1).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresDetalleO.rpt"
                        ElseIf OptRep.Item(2).Value = True Then
                            GNombreReporte = "EvaluacionProveedoresCertificadoO.rpt"
                        End If
                    End If
                                    
                    FrmReporte.Show
                
                    
MousePointer = 0
        
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BProveedor = True Then
            TxtPro.Text = DBGridBusqueda.Columns(0).Text
            TxtPro.SetFocus
        ElseIf BGrupo = True Then
            TxtPro.Text = DBGridBusqueda.Columns(2).Text
            TxtPro.SetFocus
        End If
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BProveedor = True Then
                TxtPro.Text = DBGridBusqueda.Columns(0).Text
                TxtPro.SetFocus
            ElseIf BGrupo = True Then
                TxtPro.Text = DBGridBusqueda.Columns(2).Text
                TxtPro.SetFocus
            End If
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub Form_Load()
        MskFecEva.Text = Date
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
        
End Sub

Private Sub MskFecEva_GotFocus()
        MskFecEva.SelStart = 0
        MskFecEva.SelLength = Len(MskFecEva.Text)
End Sub

Private Sub MskFecEva_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            LblLabel.Caption = "Codigo Proveedor"
        ElseIf Index = 1 Then
            LblLabel.Caption = "Codigo Grupo"
        End If
        TxtPro.SetFocus
End Sub

Private Sub Txtbusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"

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

Private Sub TxtPro_Change()
        Set RBuscaProveedor = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtPro.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtPro.Text) & "'")
            End If
            If RBuscaProveedor.RecordCount > 0 Then
                LblPro.Caption = RBuscaProveedor!Descripcion
            Else
                LblPro.Caption = ""
            End If
End Sub

Private Sub TxtPro_DblClick()
        Set RBusqueda = New ADODB.Recordset
        If OptOpcion.Item(0).Value = True Then
            BProveedor = True
            BGrupo = False
        ElseIf OptOpcion.Item(1).Value = True Then
            BProveedor = False
            BGrupo = True
        End If
        
        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "3000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus
        
End Sub

Private Sub TxtPro_GotFocus()
        TxtPro.SelStart = 0
        TxtPro.SelLength = Len(TxtPro.Text)
End Sub

Private Sub TxtPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                If OptOpcion.Item(0).Value = True Then
                    BProveedor = True
                    BGrupo = False
                ElseIf OptOpcion.Item(1).Value = True Then
                    BProveedor = False
                    BGrupo = True
                End If
                
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion, Grupo From Proveedores")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If
        
End Sub
