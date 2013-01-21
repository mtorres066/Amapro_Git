VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LiberacionEntradasProductoTerminado 
   BackColor       =   &H00C00000&
   Caption         =   "Liberacion De Entradas De Producto Terminado X Evadeva"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "LiberacionEntradasProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   285
      Left            =   3600
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
      Bindings        =   "LiberacionEntradasProductoTerminado.frx":030A
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "LiberacionEntradasProductoTerminado.frx":0326
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
      Picture         =   "LiberacionEntradasProductoTerminado.frx":0D01
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
      Picture         =   "LiberacionEntradasProductoTerminado.frx":2D73
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
      Left            =   3600
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   3
      Left            =   2160
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   2
      Left            =   2145
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "LiberacionEntradasProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RDetalleEntradasProductoTerminado As Recordset
Dim REncabezadoEntradasProductoTerminado As Recordset
Dim RBuscaEncabezadoEntradas As Recordset
Dim RBuscaCorrelativo As Recordset
Dim RBuscaPedido As Recordset
Dim RBuscaProducto As Recordset
Dim RBuscaDocumento As Recordset

Dim VFechaEntrada As Date
Dim VNumeroPedido As Double
Dim VCantidadEntrada As Double
Dim VBodega As String
Dim VProducto As String
Dim VDiasDeAtraso As Long
Dim VCorrelativo As Double

Dim mensaje As String

Private Sub CmdLiberar_Click()
                                    'VERIFICA SI ES NUMERICO
                                    If Not IsNumeric(TxtDoc.Text) Then
                                        MsgBox "El Numero De Documento Debe Ser Numerico", vbOKOnly + vbExclamation, "Verificacion"
                                        TxtDoc.SetFocus
                                        Exit Sub
                                    End If
                                    
                                    'BUSCA EL ESTADO DE EL DOCUMENTO
                                    Set RBuscaDocumento = Db.OpenRecordset("Select Estado From EncabezadoEntradasProductoTerm Where Documento = " & TxtDoc.Text)
                                    If RBuscaDocumento.RecordCount > 0 Then
                                            If RBuscaDocumento!Estado = "LIBERADO" Then
                                                MsgBox "Esta Documento Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                                                TxtDoc.SetFocus
                                                Exit Sub
                                            End If
                                    Else
                                            MsgBox "Numero De Documento No Existe", vbOKOnly + vbExclamation, "Informacion"
                                            TxtDoc.SetFocus
                                            Exit Sub
                                    End If
                                    
                                    
                                    'PREGUNTA SI QUIERE LIBERAR
                                    mensaje = MsgBox("Está Seguro Liberar El Documento " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
                                    'SI DICE QUE NO SE SALE
                                    If mensaje = vbCancel Then
                                        Exit Sub
                                    End If
                                                    
                                    MousePointer = 11
                                    
                                    'BUSCA LA FECHA DE Documento
                                    Set RBuscaDocumento = Db.OpenRecordset("Select FechaEntrada From EncabezadoEntradasProductoTerm Where Documento = " & TxtDoc.Text)
                                    If RBuscaDocumento.RecordCount > 0 Then
                                        VFechaEntrada = RBuscaDocumento!FechaEntrada
                                    Else
                                        VFechaEntrada = Date
                                    End If
                                                                                                                    
                                    'BUSCA EL ENCABEZADO DE LA ENTRADA PARA CAMBIAR EL ESTADO POR LIBERADO
                                    Set REncabezadoEntradasProductoTerminado = Db.OpenRecordset("Select * from EncabezadoEntradasProductoTerm Where Documento = " & TxtDoc.Text)
                                    If REncabezadoEntradasProductoTerminado.RecordCount > 0 Then
                                        'EDITA EL REGISTRO
                                        REncabezadoEntradasProductoTerminado.Edit
                                            REncabezadoEntradasProductoTerminado!Estado = "LIBERADO"
                                            'ASIGNA EL USUARIO QUE SUPERVISO LA LIBERACION
                                            REncabezadoEntradasProductoTerminado!Liberado = GUsuario
                                        'GRABA EL REGISTRO
                                        REncabezadoEntradasProductoTerminado.Update
                                    End If
                                            
                                    MousePointer = 0
                                    
                                    MsgBox "Documento Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
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
            DataRecepcion.RecordSource = "Select DE.Bodega, DE.FichaTecnica, FT.Descrip, DE.Cantidad, DE.Tarima, DE.FechaProduccion, DE.Linea, DE.Calidad, DE.Batch From DetalleEntradasProductoTermina as DE, FichaTecnica as FT Where DE.FichaTecnica = FT.Esp_Tec And DE.Documento = " & TxtDoc.Text & " Order By DE.FechaProduccion, DE.Linea, DE.Tarima"
            DataRecepcion.Refresh
            DBGridRecepcion.Refresh
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE ENTRADAS
            Set RBuscaEncabezadoEntradas = Db.OpenRecordset("Select FechaEntrada, Requerido, Observaciones From EncabezadoEntradasProductoTerm Where Documento = " & TxtDoc.Text)
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
      Else
             If IsNumeric(TxtDoc.Text) Then
                DataRecepcion.RecordSource = "Select DE.Bodega, DE.FichaTecnica, FT.Descrip, DE.Cantidad, DE.Tarima, DE.FechaProduccion, DE.Linea, DE.Calidad, DE.Batch From DetalleEntradasProductoTermina as DE, FichaTecnica as FT Where DE.FichaTecnica = FT.Esp_Tec And DE.Documento = " & TxtDoc.Text & " Order By DE.FechaProduccion, DE.Linea, DE.Tarima"
                DataRecepcion.Refresh
                DBGridRecepcion.Refresh
                AnchoColumnas
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
        DBGridRecepcion.Columns(0).Width = "500"
        DBGridRecepcion.Columns(1).Width = "1400"
        DBGridRecepcion.Columns(2).Width = "4000"
        DBGridRecepcion.Columns(3).Width = "700"
        DBGridRecepcion.Columns(4).Width = "500"
        DBGridRecepcion.Columns(5).Width = "1000"
        DBGridRecepcion.Columns(6).Width = "500"
        DBGridRecepcion.Columns(7).Width = "500"
End Sub
