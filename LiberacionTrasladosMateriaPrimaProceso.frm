VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LiberacionTrasladosMateriaPrimaProceso 
   BackColor       =   &H00008000&
   Caption         =   "Liberacion De Traslados De Materia Prima"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "LiberacionTrasladosMateriaPrimaProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5895
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Data DataTraslados 
      Caption         =   "Detalle Traslados"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGridTraslados 
      Bindings        =   "LiberacionTrasladosMateriaPrimaProceso.frx":08CA
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "LiberacionTrasladosMateriaPrimaProceso.frx":08E6
      TabIndex        =   4
      Top             =   2040
      Width           =   11775
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Picture         =   "LiberacionTrasladosMateriaPrimaProceso.frx":12B9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton CmdLiberar 
      Caption         =   "&Liberar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Picture         =   "LiberacionTrasladosMateriaPrimaProceso.frx":332B
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
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Observaciones"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Requerido Por"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Fecha "
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
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   600
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3780
   End
End
Attribute VB_Name = "LiberacionTrasladosMateriaPrimaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaEncabezadoTraslados As Recordset
Dim RBuscaEntradasMateriaPrima As Recordset
Dim RBuscaDetalleTraslados As Recordset
Dim RBuscaInventario As Recordset
Dim mensaje As String
Dim RBuscaCuerposPorLamina As Recordset
Dim VPesoPorCuerpo As Double


Private Sub CmdLiberar_Click()


    'BUSCA EL TRASLADO SI LO ENCUENTRA REVISA SI YA FUE LIBERADO
    Set RBuscaEncabezadoTraslados = Db.OpenRecordset("Select * From EncabezadoTrasladosMateriaPrim Where Documento = " & TxtDoc.Text)
    If RBuscaEncabezadoTraslados.RecordCount > 0 Then
            If RBuscaEncabezadoTraslados!Estado = "LIBERADO" Then
                MsgBox "Este Traslado Ya Fue Liberado", vbOKOnly + vbInformation, "Informacion"
                TxtDoc.SetFocus
                Exit Sub
            End If
    Else
        MsgBox "Traslado No Existe", vbOKOnly + vbExclamation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
    
    'PREGUNTA SI QUIERE SUPERVISAR
    mensaje = MsgBox("Está Seguro Liberar La Transaccion " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
    'SI DICE QUE NO SE SALE
    If mensaje = vbCancel Then
        Exit Sub
    End If
    
    
    'SELECCIONA TODO EL DETALLE DEL TRASLADOS
    Set RBuscaDetalleTraslados = Db.OpenRecordset("Select * From DetalleTrasladosMateriaPrimaP Where Documento = " & TxtDoc.Text)
    If RBuscaDetalleTraslados.RecordCount > 0 Then
    MousePointer = 11
    
            
            'CREA UN CICLO PARA VERIFICAR TODO EL DETALLE
            Do Until RBuscaDetalleTraslados.EOF
            
                            
                        'BUSCA EL DETALLE DE ENTRADAS MATERIA PRIMA Y MODIFICA LA BODEGA DISPONIBILIDAD PARA SABER DONDE QUEDO EL BULTO
                        Set RBuscaEntradasMateriaPrima = Db.OpenRecordset("Select BodegaDisponibilidad, CantidadTraslado, SaldoDisponibilidad, CantidadSalida, PesoEntrada, Cantidad, Peso From DetalleEntradasMateriaPrima Where NumeroIngreso = " & RBuscaDetalleTraslados!NumeroIngreso & " And Codigo = '" & RBuscaDetalleTraslados!CodigoSalida & "'")
                            If RBuscaEntradasMateriaPrima.RecordCount > 0 Then
                            
                                    'BUSCA CUANTO PESA CADA CUERPO
                                    'PESO ENTRADA DIVIDIDO EN LA CANTIDAD DE LAMINAS DIVIDIDO ENTRE LOS CUERPOS QUE TIENE LA LAMINA
                                    VPesoPorCuerpo = (RBuscaEntradasMateriaPrima!PesoEntrada / RBuscaEntradasMateriaPrima!Cantidad)
                                                                        
                                    'EDITA EL REGISTRO Y GRABA LA BODEGA DONDE QUEDO
                                    RBuscaEntradasMateriaPrima.Edit
                                        RBuscaEntradasMateriaPrima!BodegaDisponibilidad = RBuscaDetalleTraslados!BodegaEntrada
                                        RBuscaEntradasMateriaPrima!CantidadTraslado = RBuscaDetalleTraslados!CantidadReal
                                        RBuscaEntradasMateriaPrima!SaldoDisponibilidad = RBuscaDetalleTraslados!CantidadReal
                                        RBuscaEntradasMateriaPrima!CantidadSalida = 0
                                        RBuscaEntradasMateriaPrima!PESO = RBuscaDetalleTraslados!CantidadReal * VPesoPorCuerpo
                                    RBuscaEntradasMateriaPrima.Update
                            End If
                        
                        
                'AVANZA AL SIGUIENTE REGISTRO DEL DETALLE DE TRASLADO
                RBuscaDetalleTraslados.MoveNext
            Loop
            
            'MODIFICA EL ESTADO DEL TRASLADO
                RBuscaEncabezadoTraslados.Edit
                    RBuscaEncabezadoTraslados!Estado = "LIBERADO"
                    RBuscaEncabezadoTraslados!Liberado = GUsuario
                RBuscaEncabezadoTraslados.Update
                
            MousePointer = 0
            MsgBox "Traslado A Proceso Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
    End If
    

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        DataTraslados.ConnectionString = GTipoProveedor
        DataTraslados.Refresh
End Sub

Private Sub TxtDoc_Change()
On Error Resume Next
    If IsNumeric(TxtDoc.Text) Then
            DataTraslados.RecordSource = "Select DT.NumeroIngreso, DT.CodigoSalida, I.Descripcion, DT.CantidadSalida, DT.BodegaEntrada, DT.DiferenciaReqCorMas, DT.DiferenciaReqCor, DT.CantidadDesperdicio, DT.CantidadDesperdicioProveedor, DT.CantidadReal From DetalleTrasladosMateriaPrimaP as DT, CorrelativosMateriaPrima as I Where DT.CodigoSalida = I.CodigoMateriaPrima And Documento = " & TxtDoc.Text & " Order By CodigoSalida"
            DataTraslados.Refresh
            DBGridTraslados.Refresh
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE TRASLADOS
            Set RBuscaEncabezadoTraslados = Db.OpenRecordset("Select Fecha, Requerido, Observaciones From EncabezadoTrasladosMateriaPrim Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoTraslados.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoTraslados!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoTraslados!fecha
                End If
                'REQUERIDO POR
                If IsNull(RBuscaEncabezadoTraslados!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = RBuscaEncabezadoTraslados!Requerido
                End If
                'OBSERVACIONES
                If IsNull(RBuscaEncabezadoTraslados!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RBuscaEncabezadoTraslados!Observaciones
                End If
            Else
                MskFec.Text = ""
                TxtReq.Text = ""
                TxtObs.Text = ""
            End If
      
    End If
                'DataTraslados.RecordSource = "Select DT.NumeroIngreso, DT.BodegaSalida, DT.CodigoSalida, I.Descripcion, I.Existencia, DT.CantidadSalida, DT.BodegaEntrada, DT.CantidadEntrada, DT.CantidadDesperdicio From DetalleTrasladosMateriaPrimaP as DT, InventarioMateriaPrima as I Where DT.CodigoSalida = I.CodigoMateriaPrima And DT.BodegaSalida = I.Bodega And Documento = 0"
                'DataTraslados.Refresh
                'DBGridTraslados.Refresh
                'AnchoColumnas
            
      
End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Sub AnchoColumnas()
            DBGridTraslados.Columns(0).Width = "700"
            DBGridTraslados.Columns(0).Caption = "# Ingreso"
            DBGridTraslados.Columns(1).Width = "800"
            DBGridTraslados.Columns(2).Width = "3000"
            DBGridTraslados.Columns(3).Width = "600"
            DBGridTraslados.Columns(3).Caption = "Can. Sal."
            DBGridTraslados.Columns(4).Width = "600"
            DBGridTraslados.Columns(4).Caption = "Bod. Ent"
            DBGridTraslados.Columns(5).Width = "600"
            DBGridTraslados.Columns(5).Caption = "Can.Mas"
            DBGridTraslados.Columns(6).Width = "600"
            DBGridTraslados.Columns(6).Caption = "Can.Men"
            DBGridTraslados.Columns(7).Width = "600"
            DBGridTraslados.Columns(7).Caption = "Des.Proc"
            DBGridTraslados.Columns(8).Width = "600"
            DBGridTraslados.Columns(8).Caption = "Des.Prov"
            DBGridTraslados.Columns(9).Width = "600"
            DBGridTraslados.Columns(9).Caption = "Can.Real"
            
            
End Sub
