VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LiberacionEgresosMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Liberacion De Salidas De Materia Prima"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "LiberacionEgresosMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataEgresos 
      Caption         =   "Egresos"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGridEgresos 
      Bindings        =   "LiberacionEgresosMateriaPrima.frx":0442
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "LiberacionEgresosMateriaPrima.frx":045C
      TabIndex        =   7
      Top             =   2280
      Width           =   11655
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Picture         =   "LiberacionEgresosMateriaPrima.frx":0E35
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton CmdLiberar 
      Caption         =   "&Liberar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Picture         =   "LiberacionEgresosMateriaPrima.frx":2EA7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
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
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Requerido Por"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaccion"
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
      Width           =   2715
   End
End
Attribute VB_Name = "LiberacionEgresosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaEncabezadoEgresos As Recordset
Dim RBuscaNumeroIngresoEntradas As Recordset
Dim RBuscaDetalleEgresos As Recordset
Dim RBuscaInventario As Recordset
Dim RBuscaPedido As Recordset
Dim RBuscaCuerposPorLamina As Recordset
Dim RBuscaEntradasMateriaPrima As Recordset
Dim RBuscaDetalleTraslados As Recordset
Dim RBuscaDetalleDevoluciones As Recordset

Dim mensaje As String
Dim VFechaSalida As Date
Dim VDiasDeAtraso As Long

Dim VNumeroPedido As Double
Dim VCantidadSalida As Double
Dim VBodega As String
Dim VMateriaPrima As String
Dim VNumeroIngreso As Double

Dim VPesoPorCuerpo As Double


Private Sub CmdLiberar_Click()
    
    'BUSCA EL EGRESO SI LO ENCUENTRA REVISA SI YA FUE LIBERADO
    Set RBuscaEncabezadoEgresos = Db.OpenRecordset("Select * From EncabezadoEgresosMateriaPrima Where Documento = " & TxtDoc.Text)
    If RBuscaEncabezadoEgresos.RecordCount > 0 Then
            If RBuscaEncabezadoEgresos!Estado = "LIBERADO" Then
                MsgBox "Este Egreso Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                TxtDoc.SetFocus
                Exit Sub
            End If
            'ASIGNA LA FECHA DE SALIDA
            VFechaSalida = RBuscaEncabezadoEgresos!fecha
    Else
        MsgBox "Transaccion No Existe", vbOKOnly + vbExclamation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
        
    'PREGUNTA SI QUIERE LIBERAR
    mensaje = MsgBox("Está Seguro Liberar La Salida " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
    'SI DICE QUE NO SE SALE
    If mensaje = vbCancel Then
        Exit Sub
    End If
    
    MousePointer = 11
    
    'SELECCIONA TODO EL DETALLE DE LA SALIDA
    Set RBuscaDetalleEgresos = Db.OpenRecordset("Select * From DetalleEgresosMateriaPrima Where Documento = " & TxtDoc.Text)
    If RBuscaDetalleEgresos.RecordCount > 0 Then

    
            Do Until RBuscaDetalleEgresos.EOF
                'REVISA SI EXISTEN TRALADOS PENDIENTES DEL BULTO QUE NO SE HAN LIBERADO
                Set RBuscaDetalleTraslados = Db.OpenRecordset("Select * From DetalleTrasladosMateriaPrimaP as DT, EncabezadoTrasladosMateriaPrim as ET Where DT.NumeroIngreso = " & RBuscaDetalleEgresos!NumeroIngreso & " And DT.CodigoSalida = '" & RBuscaDetalleEgresos!Codigo & "' And DT.Documento = ET.Documento And ET.Estado = 'NO LIBERADO'")
                    If RBuscaDetalleTraslados.RecordCount > 0 Then
                        MsgBox "El Bulto " & RBuscaDetalleEgresos!NumeroIngreso & " Codigo " & RBuscaDetalleEgresos!Codigo & " Tiene Traslados Pendientes De Liberar", vbOKOnly + vbInformation, "Despacho No Liberado"
                        MousePointer = 0
                        Exit Sub
                    End If
                RBuscaDetalleEgresos.MoveNext
            Loop
    
    '________________________________________________________________________________________________________________
    
            'SE VUELVE A UBICAR EN EL PRIMER REGISTRO PARA SEGUIR CON EL PROCESO
            RBuscaDetalleEgresos.MoveFirst
    
            Do Until RBuscaDetalleEgresos.EOF
                'REVISA SI EXISTEN DEVOLUCIONES PENDIENTES DEL BULTO QUE NO SE HAN LIBERADO
                Set RBuscaDetalleDevoluciones = Db.OpenRecordset("Select * From DetalleDevolucionesMateriaPrima as DD, EncabezadoDevolucionesMateriaPrima as ED Where DD.NumeroIngreso = " & RBuscaDetalleEgresos!NumeroIngreso & " And DD.CodigoSalida = '" & RBuscaDetalleEgresos!Codigo & "' And DD.Documento = ED.Documento And ED.Estado = 'NO LIBERADO'")
                    If RBuscaDetalleDevoluciones.RecordCount > 0 Then
                        MsgBox "El Bulto " & RBuscaDetalleEgresos!NumeroIngreso & " Codigo " & RBuscaDetalleEgresos!Codigo & " Tiene Devoluciones Pendientes De Liberar", vbOKOnly + vbInformation, "Despacho No Liberado"
                        MousePointer = 0
                        Exit Sub
                    End If
                RBuscaDetalleEgresos.MoveNext
            Loop
            
    
    
    '________________________________________________________________________________________________________________
            'SE VUELVE A UBICAR EN EL PRIMER REGISTRO PARA SEGUIR CON EL PROCESO
            RBuscaDetalleEgresos.MoveFirst
    
            'CREA UN CICLO PARA VERIFICAR TODO EL DETALLE
            Do Until RBuscaDetalleEgresos.EOF
                                
                    'CANTIDAD PARA REBAJAR DE PEDIDO
                     VCantidadSalida = RBuscaDetalleEgresos!Cantidad
                    'BODEGA PARA BUSCAR MATERIA PRIMA
                     VBodega = RBuscaDetalleEgresos!Bodega
                    'CODIGO DE MATERIA PRIMA
                     VMateriaPrima = RBuscaDetalleEgresos!Codigo
                    'NUMERO DE INGRESO MATERIA PRIMA QUE SALE
                     VNumeroIngreso = RBuscaDetalleEgresos!NumeroIngreso
            
                            
            'NUMERO INGRESO EN ENTRADAS ------------------------------------
                     'BUSCA EL NUMERO DE INGRESO CON CODIGO DE MATERIA PRIMA EN EL DETALLE DE LAS ENTRADAS
                      Set RBuscaNumeroIngresoEntradas = Db.OpenRecordset("Select SaldoDisponibilidad, CantidadSalida, PesoEntrada, CantidadTraslado, Peso From DetalleEntradasMateriaPrima Where Codigo = '" & VMateriaPrima & "' And NumeroIngreso = " & VNumeroIngreso)
                          
                          If RBuscaNumeroIngresoEntradas.RecordCount > 0 Then
                     
                            
                        'BUSCA EL DETALLE DE ENTRADAS MATERIA PRIMA
                        Set RBuscaEntradasMateriaPrima = Db.OpenRecordset("Select PesoEntrada, Cantidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & VNumeroIngreso & " And Codigo = '" & VMateriaPrima & "'")
                            If RBuscaEntradasMateriaPrima.RecordCount > 0 Then
                                    'BUSCA CUANTO PESA CADA CUERPO
                                    'PESO ENTRADA DIVIDIDO EN LA CANTIDAD DE LAMINAS DIVIDIDO ENTRE LOS CUERPOS QUE TIENE LA LAMINA
                                     VPesoPorCuerpo = RBuscaEntradasMateriaPrima!PesoEntrada / RBuscaEntradasMateriaPrima!Cantidad
                            End If
                                                 
                            'ASIGNA LA DISPONIBILIDAD AL BULTO O NUMERO DE INGRESO Y CONTROLA EL SALDO DEL BULTO
                            RBuscaNumeroIngresoEntradas.Edit
                                    RBuscaNumeroIngresoEntradas!CantidadSalida = RBuscaNumeroIngresoEntradas!CantidadSalida + VCantidadSalida
                                    RBuscaNumeroIngresoEntradas!SaldoDisponibilidad = RBuscaNumeroIngresoEntradas!SaldoDisponibilidad - VCantidadSalida
                                    'DESCONTAMOS EL PESO DE EL BULTO DE ACUERDO A LA CANTIDAD DE SALIDA POR EL PESO DE CADA CUERPO
                                    RBuscaNumeroIngresoEntradas!PESO = (RBuscaNumeroIngresoEntradas!PESO - (VCantidadSalida * VPesoPorCuerpo))
                                                                        
                            RBuscaNumeroIngresoEntradas.Update
                            If Err <> 0 Then
                            End If
                     End If
                                  
                       
                'AVANZA AL SIGUIENTE REGISTRO DEL DETALLE DE TRASLADO
                RBuscaDetalleEgresos.MoveNext
            Loop
            
            'MODIFICA EL ESTADO DEL EGRESO
                RBuscaEncabezadoEgresos.Edit
                    RBuscaEncabezadoEgresos!Estado = "LIBERADO"
                    RBuscaEncabezadoEgresos!Liberado = GUsuario
                RBuscaEncabezadoEgresos.Update
                
            MousePointer = 0
            MsgBox "Salida Liberada Con Exito", vbOKOnly + vbInformation, "Informacion"
    End If
                              
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        DataEgresos.ConnectionString = GTipoProveedor
        DataEgresos.Refresh
End Sub

Private Sub TxtDoc_Change()
On Error Resume Next
    If IsNumeric(TxtDoc.Text) Then
            DataEgresos.RecordSource = "Select DE.Bodega, DE.NumeroIngreso, DE.Codigo, C.Descripcion, DE.Cantidad From DetalleEgresosMateriaPrima as DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And DE.Documento = " & TxtDoc.Text & " Order By DE.Codigo"
            DataEgresos.Refresh
            DBGridEgresos.Refresh
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE EGRESOS
            Set RBuscaEncabezadoEgresos = Db.OpenRecordset("Select Fecha, Requerido, Observaciones From EncabezadoEgresosMateriaPrima Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoEgresos.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoEgresos!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoEgresos!fecha
                End If
                'REQUERIDO POR
                If IsNull(RBuscaEncabezadoEgresos!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = RBuscaEncabezadoEgresos!Requerido
                End If
                'OBSERVACIONES
                If IsNull(RBuscaEncabezadoEgresos!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RBuscaEncabezadoEgresos!Observaciones
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
        DBGridEgresos.Columns(0).Width = "1200"
        DBGridEgresos.Columns(1).Width = "1500"
        DBGridEgresos.Columns(2).Width = "1200"
        DBGridEgresos.Columns(3).Width = "4000"
        DBGridEgresos.Columns(4).Width = "1200"
End Sub

