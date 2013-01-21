VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LiberacionEntradasMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Liberacion De Entradas De Materia Prima"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "LiberacionEntradasMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Data DataRecepcion 
      Caption         =   "Recepcion"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGridRecepcion 
      Bindings        =   "LiberacionEntradasMateriaPrima.frx":030A
      Height          =   5655
      Left            =   120
      OleObjectBlob   =   "LiberacionEntradasMateriaPrima.frx":0326
      TabIndex        =   4
      Top             =   2280
      Width           =   11655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "LiberacionEntradasMateriaPrima.frx":0D01
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton CmdLiberar 
      Caption         =   "&Liberar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "LiberacionEntradasMateriaPrima.frx":2D73
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox TxtDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Requerido"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2745
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Transaccion"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3780
   End
End
Attribute VB_Name = "LiberacionEntradasMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RDetalleEntradasMateriaPrima As Recordset
Dim REncabezadoEntradasMateriaPrima As Recordset
Dim RBuscaEncabezadoEntradas As Recordset
Dim RBuscaCorrelativo As Recordset
Dim RBuscaPedido As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaRecepcion As Recordset

Dim VFechaEntrada As Date
Dim VNumeroPedido As String
Dim VCantidadEntrada As Double
Dim VBodega As String
Dim VMateriaPrima As String
Dim VDiasDeAtraso As Double
Dim VCorrelativo As Double

Dim mensaje As String

Private Sub CmdLiberar_Click()
                                    
                                    'BUSCA EL ESTADO DE LA RECEPCION
                                    Set RBuscaRecepcion = Db.OpenRecordset("Select Estado From EncabezadoEntradasMateriaPrima Where Documento = " & TxtDoc.Text)
                                    If RBuscaRecepcion.RecordCount > 0 Then
                                            If RBuscaRecepcion!Estado = "LIBERADO" Then
                                                MsgBox "Esta Recepcion Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
                                                TxtDoc.SetFocus
                                                Exit Sub
                                            End If
                                    Else
                                            MsgBox "Numero De Transaccion No Existe", vbOKOnly + vbExclamation, "Informacion"
                                            TxtDoc.SetFocus
                                            Exit Sub
                                    End If
                                    
                                    'PREGUNTA SI QUIERE SUPERVISAR
                                    mensaje = MsgBox("Está Seguro Liberar La Recepcion " & TxtDoc.Text, vbOKCancel + vbInformation + vbDefaultButton2, "Verificacion")
                                    
                                    'SI DICE QUE NO SE SALE
                                    If mensaje = vbCancel Then
                                        Exit Sub
                                    End If
                                                    
                                    'BUSCA SI HAY BULTOS NO INSPECCIONADOS
                                    Set RDetalleEntradasMateriaPrima = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Documento = " & TxtDoc.Text & " And Estado = 'N'")
                                        If RDetalleEntradasMateriaPrima.RecordCount > 0 Then
                                            MsgBox "Esta Transaccion No Se Puede Liberar, Uno o Mas Bultos No Han Sido Inspeccionadas", vbOKOnly + vbExclamation, "Verifique"
                                            Exit Sub
                                        End If
                                                 
                                    MousePointer = 11
                                    
                                    'BUSCA LA FECHA DE RECEPCION
                                    'Set RBuscaRecepcion = Db.OpenRecordset("Select FechaEntrada From EncabezadoEntradasMateriaPrima Where Documento = '" & TxtDoc.Text & "'")
                                    'If RBuscaRecepcion.RecordCount > 0 Then
                                    '    VFechaEntrada = RBuscaRecepcion!FechaEntrada
                                    'Else
                                    '    VFechaEntrada = Date
                                    'End If
                                    
                                    'SELECCIONAMOS TODOS LOS DETALLE DE LA ENTRADA DE ACUERDO AL NUMERO DE RECEPCION
                                    'ORDENADO POR FECHA DE PRODUCCION DE PROVEEDOR
                                    Set RDetalleEntradasMateriaPrima = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Documento = " & TxtDoc.Text & " Order By Codigo, FechaBoleta, BultoBoleta")
                                                                            
                                    'SE POSICIONA EN EL PRIMER REGISTRO DEL RECORDSET
                                    RDetalleEntradasMateriaPrima.MoveFirst
                                    
                                    'EJECUTA UN CICLO HASTA EL FINAL PARA PODER CALCULAR LOS CODIGOS DE INGRESO
                                    'Y MODIFICAR EL PEDIDO Y LA EXISTENCIA DE MATERIA PRIMA
                                    Do Until RDetalleEntradasMateriaPrima.EOF
                                    
                                            'NUMERO DE PEDIDO
                                            'VNumeroPedido = RDetalleEntradasMateriaPrima!DocumentoPedido
                                            'CANTIDAD PARA REBAJAR DE PEDIDO
                                            VCantidadEntrada = RDetalleEntradasMateriaPrima!Cantidad
                                            'BODEGA PARA BUSCAR MATERIA PRIMA
                                            VBodega = RDetalleEntradasMateriaPrima!BodegaDisponibilidad
                                            'CODIGO DE MATERIA PRIMA
                                            VMateriaPrima = RDetalleEntradasMateriaPrima!Codigo
                                                                                
                                            'BUSCA EL CORRELATIVO DE INGRESO MAXIMO DE ACUERDO AL CODIGO DE MATERIA PRIMA
                                            'Y LE AUMENTA 1 Y SE LO ASIGNA AL NUEVO CODIGO DE MATERIA PRIMA
                                            Set RBuscaCorrelativo = Db.OpenRecordset("Select Correlativo From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & VMateriaPrima & "'")
                                            'SI LO ENCUENTRA
                                            If RBuscaCorrelativo.RecordCount > 0 Then
                                                'MODIFICA EL CORRELATIVO
                                                RBuscaCorrelativo.Edit
                                                    RBuscaCorrelativo!Correlativo = Val(RBuscaCorrelativo!Correlativo) + 1
                                                    VCorrelativo = RBuscaCorrelativo!Correlativo
                                                RBuscaCorrelativo.Update
                                                
                                                'SI HAY ERROR NO HACE NADA
                                                If Err <> 0 Then
                                                End If
                                            Else
                                                'AGREGA EL CODIGO DE MATERIA PRIMA Y EL CORRELATIVO
                                                RBuscaCorrelativo.AddNew
                                                    RBuscaCorrelativo!CodigoMateriaPrima = VMateriaPrima
                                                    RBuscaCorrelativo!Correlativo = "1"
                                                RBuscaCorrelativo.Update
                                                
                                                'SI HAY ERROR NO HACE NADA
                                                If Err <> 0 Then
                                                End If
                                                
                                                VCorrelativo = "1"
                                            End If
                                                
                                            'ASIGNA EL CORRELATIVO Y EL ESTADO DEL DETALLE
                                            RDetalleEntradasMateriaPrima.Edit
                                                RDetalleEntradasMateriaPrima!NumeroIngreso = VCorrelativo
                                                RDetalleEntradasMateriaPrima!Barra = VMateriaPrima & "-" & VCorrelativo
                                            RDetalleEntradasMateriaPrima.Update
                                                
                                            'SI HAY ERROR NO HACE NADA
                                            If Err <> 0 Then
                                            End If
                                            
                                            '---------- PEDIDO ------------------------------------------------------
                                                        
                                            'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                                            'Set RBuscaPedido = Db.OpenRecordset("Select CantidadEntregada, SaldoPorEntregar, FechaParaEntregar, FechaEntregaTotal, DiasDeAtraso From PedidosMateriaPrimaProveedores Where Documento = '" & VNumeroPedido & "'")
                                            '    If RBuscaPedido.RecordCount > 0 Then
                                            '        'EDITA EL REGISTRO DE PEDIDO Y ACTUALIZA DATOS
                                            '         RBuscaPedido.Edit
                                            '                 RBuscaPedido!CantidadEntregada = RBuscaPedido!CantidadEntregada + VCantidadEntrada
                                            '                 RBuscaPedido!SaldoPorEntregar = RBuscaPedido!SaldoPorEntregar - VCantidadEntrada
                                            '
                                            '                'SI EL SALDO POR ENTREGAR YA ESTA EN CERO O MENOR QUE CERO ACTUALIZA LA FECHA DE ENTREGA Y CALCULA
                                            '                'LOS DIAS DE ATRASO
                                            '                If RBuscaPedido!SaldoPorEntregar <= 0 Then
                                            '                    'CAMBIA LA FECHA DE ENTREGA TOTAL POR LA ACTUAL DEL ULTIMO INGRESO
                                            '                    RBuscaPedido!FechaEntregaTotal = VFechaEntrada
                                            '
                                            '                    'CALCULA LOS DIAS DE ATRASO
                                            '                    VDiasDeAtraso = DateValue(VFechaEntrada) - DateValue(RBuscaPedido!FechaParaEntregar)
                                            '
                                            '                    'SI LA VARIABLE VDIASDEATRASO ES MENOR QUE CERO ES PORQUE ENTREGO EL PEDIDO ANTES DE LA FECHA
                                            '                    If VDiasDeAtraso < 0 Then
                                            '                        VDiasDeAtraso = 0
                                            '                    End If
                                            '
                                            '                    'MODIFICA LOS DIAS DE ATRASO EN EL PEDIDO
                                            '                    RBuscaPedido!DiasDeAtraso = VDiasDeAtraso
                                            '                Else
                                            '                    If IsNull(RBuscaPedido!FechaEntregaTotal) Then
                                            '                    Else
                                            '                        RBuscaPedido!FechaEntregaTotal = ""
                                            '                        RBuscaPedido!DiasDeAtraso = "0"
                                            '                    End If
                                            '                End If
                                            '            'GRABA DATOS
                                            '            RBuscaPedido.Update
                                            '    End If
                                            
                                            'SE MUEVE AL SIGUIENTE REGISTRO
                                            RDetalleEntradasMateriaPrima.MoveNext
                                            
                                    Loop 'FIN DE CICLO
                                                                                
                                                                                
                                    '-----------------------  BULTO -------------------------------------------------------------------------------
                                    
                                    Set REncabezadoEntradasMateriaPrima = Db.OpenRecordset("Select * from EncabezadoEntradasMateriaPrima Where Documento = " & TxtDoc.Text)
                                    If REncabezadoEntradasMateriaPrima.RecordCount > 0 Then
                                        'EDITA EL REGISTRO
                                        REncabezadoEntradasMateriaPrima.Edit
                                            REncabezadoEntradasMateriaPrima!Estado = "LIBERADO"
                                            'ASIGNA EL USUARIO QUE SUPERVISO LA LIBERACION
                                            REncabezadoEntradasMateriaPrima!Liberado = GUsuario
                                        'GRABA EL REGISTRO
                                        REncabezadoEntradasMateriaPrima.Update
                                    End If
                                            
                                    MousePointer = 0
                                    
                                    MsgBox "Recepcion Liberada Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
            DataRecepcion.ConnectionString = GTipoProveedor
            DataRecepcion.Refresh
End Sub

Private Sub TxtDoc_Change()
On Error Resume Next
    If IsNumeric(TxtDoc.Text) Then
            DataRecepcion.RecordSource = "Select DE.NumeroIngreso, DE.Codigo, I.Descripcion, DE.Cantidad, DE.Peso, DE.Calidad From DetalleEntradasMateriaPrima as DE, CorrelativosMateriaPrima as I Where DE.Codigo = I.CodigoMateriaPrima And DE.Documento = " & TxtDoc.Text & " Order By DE.BodegaDisponibilidad, DE.Codigo"
            DataRecepcion.Refresh
            DBGridRecepcion.Refresh
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE ENTRADAS
            Set RBuscaEncabezadoEntradas = Db.OpenRecordset("Select FechaEntrada, Requerido, Observaciones From EncabezadoEntradasMateriaPrima Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoEntradas.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoEntradas!FechaEntrada) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoEntradas!FechaEntrada
                End If
                'REQUERIDO POR
                If IsNull(RBuscaEncabezadoEntradas!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = RBuscaEncabezadoEntradas!Requerido
                End If
                'OBSERVACIONES
                If IsNull(RBuscaEncabezadoEntradas!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RBuscaEncabezadoEntradas!Observaciones
                End If
            Else
                MskFec.Text = ""
                TxtReq.Text = ""
                TxtObs.Text = ""
            End If
    End If
End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Sub AnchoColumnas()
        DBGridRecepcion.Columns(0).Width = "1100"
        DBGridRecepcion.Columns(0).Caption = "# Ingreso"
        DBGridRecepcion.Columns(1).Width = "1100"
        DBGridRecepcion.Columns(2).Width = "4000"
        DBGridRecepcion.Columns(3).Width = "1000"
        DBGridRecepcion.Columns(4).Width = "1000"
        DBGridRecepcion.Columns(4).Caption = "Peso"
        DBGridRecepcion.Columns(5).Width = "1000"
        
        

End Sub
