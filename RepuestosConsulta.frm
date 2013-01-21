VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RepuestosConsulta 
   Caption         =   "Consulta De Repuestos"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "RepuestosConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Buscar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   10920
         Picture         =   "RepuestosConsulta.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   5415
      End
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   7335
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
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
   End
   Begin VB.Frame FrameProductos2 
      Caption         =   "Tipo De Existencia"
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
      Height          =   1095
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton OptExi 
         Caption         =   "Mayor Que Cero"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptExi 
         Caption         =   "Cero"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptExi 
         Caption         =   "Todos"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptExi 
         Caption         =   "Menor Que Cero"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   11040
      Picture         =   "RepuestosConsulta.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   10200
      Picture         =   "RepuestosConsulta.frx":43EE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Consultar"
      Top             =   120
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "click en encabezado columna para indexar"
      Top             =   1560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12091
      _Version        =   393216
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
   Begin VB.TextBox TxtTex 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones De Busqueda"
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
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton OptOpc 
         Caption         =   "Maquina"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Ubicacion"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Clasificacion"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Tipo De Producto"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
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
      Left            =   8280
      TabIndex        =   8
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label LblEti 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "RepuestosConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaRepuesto As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RConsultar As New ADODB.Recordset
Dim RBuscaDatos As New ADODB.Recordset
Dim RBusca As New ADODB.Recordset

Dim BCodigo As Boolean
Dim BDescripcion As Boolean
Dim BTipo As Boolean
Dim BClasificacion As Boolean
Dim BUbicacion As Boolean
Dim BMaquina As Boolean

Dim VCriteria As String

Private Sub Command1_Click()

        
        'CODIGO
        If OptOpc.Item(0).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where I.Codigo Like '%" & TxtTex.Text & "%' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(I.Codigo) Like '%" & UCase(TxtTex.Text) & "%' And UPPER(I.Codigo) = UPPER(E.Codigo) And UPPER(I.Unidadmedida) = UPPER(U.Codigo)"
            End If
        'DESCRIPCION
        ElseIf OptOpc.Item(1).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where I.Descripcion Like '%" & TxtTex.Text & "%' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(I.Descripcion) Like '%" & UCase(TxtTex.Text) & "%' And UPPER(I.Codigo) = UPPER(E.Codigo) And UPPER(I.Unidadmedida) = UPPER(U.Codigo)"
            End If
        'TIPO DE PRODUCTO
        ElseIf OptOpc.Item(2).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where I.TipoProducto Like '%" & TxtTex.Text & "%' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(I.TipoProducto) Like '%" & UCase(TxtTex.Text) & "%' And UPPER(I.Codigo) = UPPER(E.Codigo) And UPPER(I.Unidadmedida) = UPPER(U.Codigo)"
            End If
        'CLASIFICACION
        ElseIf OptOpc.Item(3).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where I.Clasificacion Like '%" & TxtTex.Text & "%' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(I.Clasificacion) Like '%" & UCase(TxtTex.Text) & "%' And UPPER(I.Codigo) = UPPER(E.Codigo) And UPPER(I.Unidadmedida) = UPPER(U.Codigo)"
            End If
        'UBICACION
        ElseIf OptOpc.Item(4).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where E.Ubicacion = '" & TxtTex.Text & "' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(E.Ubicacion) = '" & UCase(TxtTex.Text) & "' And UPPER(I.Codigo) = UPPER(E.Codigo And UPPER(I.Unidadmedida) = UPPER(U.Codigo))"
            End If
        'MAQUINA
        ElseIf OptOpc.Item(5).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where I.Maquina Like '%" & TxtTex.Text & "%' And I.Codigo = E.Codigo And I.UnidadMedida = U.Codigo"
            Else
                VCriteria = "Select I.Codigo, I.Descripcion, E.Ubicacion, E.Existencia, U.Descripcion, I.CostoPromedio, (E.Existencia * I.CostoPromedio) From M_Repuestos I, M_RepuestosInventario E, M_RepuestosUnidadMedida U Where UPPER(I.Maquina) Like '%" & UCase(TxtTex.Text) & "%' And UPPER(I.Codigo) = UPPER(E.Codigo) And UPPER(I.Unidadmedida) = UPPER(U.Codigo)"
            End If
        End If
        
        'MAYOR QUE CERO
        If OptExi.Item(0).Value = True Then
            VCriteria = VCriteria & " And E.Existencia > 0"
        'IGUAL A CERO
        ElseIf OptExi.Item(1).Value = True Then
            VCriteria = VCriteria & " And E.Existencia = 0"
        'MENOR QUE CERO
        ElseIf OptExi.Item(2).Value = True Then
            VCriteria = VCriteria & " And E.Existencia < 0"
        'TODOS
        ElseIf OptExi.Item(3).Value = True Then
        
        End If
        
        
        
        Set RConsultar = New ADODB.Recordset
        Call Abrir_Recordset(RConsultar, VCriteria)
        Set DataGrid1.DataSource = RConsultar
        DataGrid1.Columns(0).Width = 1100
        DataGrid1.Columns(1).Width = 4000
        DataGrid1.Columns(2).Width = 1800
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Width = 800
        DataGrid1.Columns(5).Width = 1000
        DataGrid1.Columns(6).Width = 1300
        
        
        DataGrid1.Columns(3).NumberFormat = ("#,###,##0.00")
        DataGrid1.Columns(3).Alignment = dbgRight
        DataGrid1.Columns(5).NumberFormat = ("#,###,##0.00")
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(6).NumberFormat = ("#,###,##0.00")
        DataGrid1.Columns(6).Alignment = dbgRight
        
        DataGrid1.Columns(0).Caption = "Codigo"
        DataGrid1.Columns(1).Caption = "Descripcion"
        DataGrid1.Columns(2).Caption = "Ubicacion"
        DataGrid1.Columns(3).Caption = "Existencia"
        DataGrid1.Columns(4).Caption = "U.Medida"
        DataGrid1.Columns(5).Caption = "Costo Promedio"
        DataGrid1.Columns(6).Caption = "Costo Total"
            
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
        On Error Resume Next
                RConsultar.Sort = RConsultar.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

End Sub

Private Sub DBGridBusqueda_DblClick()
        
        If BCodigo = True Or BTipo = True Or BClasificacion = True Or BUbicacion = True Or BMaquina = True Then
            TxtTex.Text = DbGridBusqueda.Columns(0).Text
            TxtTex.SetFocus
        ElseIf BDescripcion = True Then
            TxtTex.Text = DbGridBusqueda.Columns(1).Text
            TxtTex.SetFocus
        End If
        FrameBuscar.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BCodigo = True Or BTipo = True Or BClasificacion = True Or BUbicacion = True Or BMaquina = True Then
                TxtTex.Text = DbGridBusqueda.Columns(0).Text
                TxtTex.SetFocus
            ElseIf BDescripcion = True Then
                TxtTex.Text = DbGridBusqueda.Columns(1).Text
                TxtTex.SetFocus
            End If
            FrameBuscar.Visible = False
        End If
End Sub

Private Sub OptOpc_Click(Index As Integer)
    If Index = 0 Then
        LblEti.Caption = "Codigo"
    ElseIf Index = 1 Then
        LblEti.Caption = "Descripcion"
    ElseIf Index = 2 Then
        LblEti.Caption = "Tipo De Producto"
    ElseIf Index = 3 Then
        LblEti.Caption = "Clasificacion"
    ElseIf Index = 4 Then
        LblEti.Caption = "Ubicacion"
     ElseIf Index = 5 Then
        LblEti.Caption = "Maquina"
    End If
        TxtTex.SetFocus
End Sub

Private Sub Txtbuscar_Change()
            Set RBusqueda = New ADODB.Recordset

            
            If BCodigo = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                        If OptCodigo.Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where Codigo Like '%" & TxtBuscar.Text & "%'")
                        ElseIf OptDescripcion.Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                        End If
                    Else
                        If OptCodigo.Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                        ElseIf OptDescripcion.Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                        End If
                    End If
            ElseIf BDescripcion = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where Codigo Like '%" & TxtBuscar.Text & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                            End If
                    Else
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            'TIPO DE PRODUCTO
            ElseIf BTipo = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto Where Codigo Like '%" & TxtBuscar.Text & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                            End If
                    Else
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            'CLASIFICACION
            ElseIf BClasificacion = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion Where Codigo Like '%" & TxtBuscar.Text & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                            End If
                    Else
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            'CLASIFICACION
            ElseIf BUbicacion = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones Where Codigo Like '%" & TxtBuscar.Text & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                            End If
                    Else
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            'MAQUINA
            ElseIf BMaquina = True Then
                    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas Where Codigo Like '%" & TxtBuscar.Text & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                            End If
                    Else
                            If OptCodigo.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            ElseIf OptDescripcion.Value = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            End If
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBuscar_GotFocus()
        TxtBuscar.SelStart = 0
        TxtBuscar.SelLength = Len(TxtBuscar.Text)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTex_Change()
        'CODIGO
        If OptOpc.Item(0).Value = True Then
                Set RBusca = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_Repuestos Where Codigo = '" & TxtTex.Text & "'")
                    Else
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_Repuestos Where UPPER(Codigo) = '" & UCase(TxtTex.Text) & "'")
                    End If
                    If RBusca.RecordCount > 0 Then
                        LblDes.Caption = RBusca!Descripcion
                    Else
                        LblDes.Caption = ""
                    End If
        'TIPO
        ElseIf OptOpc.Item(2).Value = True Then
                Set RBusca = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_RepuestosTiposProducto Where Codigo = '" & TxtTex.Text & "'")
                    Else
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_RepuestosTiposProducto Where UPPER(Codigo) = '" & UCase(TxtTex.Text) & "'")
                    End If
                    If RBusca.RecordCount > 0 Then
                        LblDes.Caption = RBusca!Descripcion
                    Else
                        LblDes.Caption = ""
                    End If
        'CLASIFICACION
        ElseIf OptOpc.Item(3).Value = True Then
                Set RBusca = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_RepuestosClasificacion Where Codigo = '" & TxtTex.Text & "'")
                    Else
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_RepuestosClasificacion Where UPPER(Codigo) = '" & UCase(TxtTex.Text) & "'")
                    End If
                    If RBusca.RecordCount > 0 Then
                        LblDes.Caption = RBusca!Descripcion
                    Else
                        LblDes.Caption = ""
                    End If
        
        'MAQUINA
        ElseIf OptOpc.Item(5).Value = True Then
                Set RBusca = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_Maquinas Where Codigo = '" & TxtTex.Text & "'")
                    Else
                        Call Abrir_Recordset(RBusca, "Select Descripcion From M_Maquinas Where UPPER(Codigo) = '" & UCase(TxtTex.Text) & "'")
                    End If
                    If RBusca.RecordCount > 0 Then
                        LblDes.Caption = RBusca!Descripcion
                    Else
                        LblDes.Caption = ""
                    End If
        
        End If

End Sub

Private Sub TxtTex_DblClick()
            Set RBusqueda = New ADODB.Recordset
                'CODIGO
                If OptOpc.Item(0).Value = True Then
                             BCodigo = True
                             BDescripcion = False
                             BTipo = False
                             BClasificacion = False
                             BUbicacion = False
                             BMaquina = False
                             Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos")
                 'DESCRIPCOIN
                 ElseIf OptOpc.Item(1).Value = True Then
                            BCodigo = False
                            BDescripcion = True
                            BTipo = False
                            BClasificacion = False
                            BUbicacion = False
                            BMaquina = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos")
                 'TIPO DE PRODUCTO
                 ElseIf OptOpc.Item(2).Value = True Then
                            BCodigo = False
                            BDescripcion = False
                            BTipo = True
                            BClasificacion = False
                            BUbicacion = False
                            BMaquina = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto")
                'CLASIFICACION
                ElseIf OptOpc.Item(3).Value = True Then
                            BCodigo = False
                            BDescripcion = False
                            BTipo = False
                            BClasificacion = True
                            BUbicacion = False
                            BMaquina = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion")
                'UBICACIONES
                ElseIf OptOpc.Item(4).Value = True Then
                            BCodigo = False
                            BDescripcion = False
                            BTipo = False
                            BClasificacion = False
                            BUbicacion = True
                            BMaquina = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones")
                'MAQUINAS
                ElseIf OptOpc.Item(5).Value = True Then
                            BCodigo = False
                            BDescripcion = False
                            BTipo = False
                            BClasificacion = False
                            BUbicacion = False
                            BMaquina = True
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas")
                End If
                            FrameBuscar.Visible = True
                            TxtBuscar.SetFocus
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtTex_GotFocus()
        TxtTex.SelStart = 0
        TxtTex.SelLength = Len(TxtTex.Text)
End Sub

Private Sub TxtTex_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                                            'CODIGO
                                            If OptOpc.Item(0).Value = True Then
                                                         BCodigo = True
                                                         BDescripcion = False
                                                         BTipo = False
                                                         BClasificacion = False
                                                         BUbicacion = False
                                                         BMaquina = False
                                                         Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos")
                                             'DESCRIPCOIN
                                             ElseIf OptOpc.Item(1).Value = True Then
                                                        BCodigo = False
                                                        BDescripcion = True
                                                        BTipo = False
                                                        BClasificacion = False
                                                        BUbicacion = False
                                                        BMaquina = False
                                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Repuestos")
                                             'TIPO DE PRODUCTO
                                             ElseIf OptOpc.Item(2).Value = True Then
                                                        BCodigo = False
                                                        BDescripcion = False
                                                        BTipo = True
                                                        BClasificacion = False
                                                        BUbicacion = False
                                                        BMaquina = False
                                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosTiposProducto")
                                            'CLASIFICACION
                                            ElseIf OptOpc.Item(3).Value = True Then
                                                        BCodigo = False
                                                        BDescripcion = False
                                                        BTipo = False
                                                        BClasificacion = True
                                                        BUbicacion = False
                                                        BMaquina = False
                                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosClasificacion")
                                            'UBICACIONES
                                            ElseIf OptOpc.Item(4).Value = True Then
                                                        BCodigo = False
                                                        BDescripcion = False
                                                        BTipo = False
                                                        BClasificacion = False
                                                        BUbicacion = True
                                                        BMaquina = False
                                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_RepuestosUbicaciones")
                                            'MAQUINAS
                                            ElseIf OptOpc.Item(5).Value = True Then
                                                        BCodigo = False
                                                        BDescripcion = False
                                                        BTipo = False
                                                        BClasificacion = False
                                                        BUbicacion = False
                                                        BMaquina = True
                                                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from M_Maquinas")
                                            End If
                                            FrameBuscar.Visible = True
                            TxtBuscar.SetFocus
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "4000"
        End If
        
End Sub
