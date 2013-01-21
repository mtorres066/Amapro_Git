VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form LiberacionIngresosMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Liberacion De INGRESOS De Materia Prima"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "LiberacionIngresosMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataIngresos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGridIngresos 
      Bindings        =   "LiberacionIngresosMateriaPrima.frx":0442
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "LiberacionIngresosMateriaPrima.frx":045D
      TabIndex        =   7
      Top             =   2040
      Width           =   11655
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1680
      Width           =   5655
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
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
      Height          =   855
      Left            =   9960
      Picture         =   "LiberacionIngresosMateriaPrima.frx":0E37
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
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
      Height          =   855
      Left            =   9960
      Picture         =   "LiberacionIngresosMateriaPrima.frx":1279
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
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
      Top             =   120
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
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
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
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Documento"
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
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3540
   End
End
Attribute VB_Name = "LiberacionIngresosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaEncabezadoIngresos As Recordset
Dim RBuscaDetalleIngresos As Recordset
Dim RBuscaInventario As Recordset
Dim Mensaje As String

Private Sub CmdLiberar_Click()
If Not IsNumeric(TxtDoc.Text) Then
        MsgBox "El Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbExclamation, "Verificacion"
        TxtDoc.SetFocus
        Exit Sub
    End If

    'BUSCA EL EGRGESO SI LO ENCUENTRA REVISA SI YA FUE LIBERADO
    Set RBuscaEncabezadoIngresos = Db.OpenRecordset("Select * From EncabezadoIngresosMateriaPrima Where Documento = " & TxtDoc.Text)
    If RBuscaEncabezadoIngresos.RecordCount > 0 Then
            If RBuscaEncabezadoIngresos!Estado = "LIBERADO" Then
                MsgBox "Este Ingreso Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                TxtDoc.SetFocus
                Exit Sub
            End If
    Else
        MsgBox "Egreso No Existe", vbOKOnly + vbExclamation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
    
    'PREGUNTA SI QUIERE SUPERVISAR
    Mensaje = MsgBox("Está Seguro Liberar El Ingreso " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
    'SI DICE QUE NO SE SALE
    If Mensaje = vbCancel Then
        Exit Sub
    End If
        
    
    'SELECCIONA TODO EL DETALLE DEL TRASLADOS
    Set RBuscaDetalleIngresos = Db.OpenRecordset("Select * From DetalleIngresosMateriaPrima Where Documento = " & TxtDoc.Text)
    If RBuscaDetalleIngresos.RecordCount > 0 Then
    MousePointer = 11
            'CREA UN CICLO PARA VERIFICAR TODO EL DETALLE
            Do Until RBuscaDetalleIngresos.EOF
                        'BUSCA EN INVENTARIO EL CODIGO DE SALIDA DEL DETALLE DE TRASLADOS Y SUMA LOS TRASLADOS Y RESTA LA EXISTENCIA
                        Set RBuscaInventario = Db.OpenRecordset("Select Entradas, Existencia From InventarioMateriaPrima Where CodigoMateriaPrima = '" & RBuscaDetalleIngresos!Codigo & "' And Bodega = '" & RBuscaDetalleIngresos!Bodega & "'")
                            'SI ENCUENTRA LA MATERIA PRIMA
                            If RBuscaInventario.RecordCount > 0 Then
                                RBuscaInventario.Edit
                                    'SUMA LA CANTIDAD DE TRASLADOS EXISTENTE MAS LA CANTIDAD QUE ESTAN TRASLADANDO
                                    RBuscaInventario!Entradas = RBuscaInventario!Entradas + RBuscaDetalleIngresos!Cantidad
                                    RBuscaInventario!Existencia = RBuscaInventario!Existencia + RBuscaDetalleIngresos!Cantidad
                                RBuscaInventario.Update
                            
                                If Err <> 0 Then
                                End If
                            End If
                       
                'AVANZA AL SIGUIENTE REGISTRO DEL DETALLE DE TRASLADO
                RBuscaDetalleIngresos.MoveNext
            Loop
            
               'MODIFICA EL ESTADO DEL INGRESO
                RBuscaEncabezadoIngresos.Edit
                    RBuscaEncabezadoIngresos!Estado = "LIBERADO"
                    RBuscaEncabezadoIngresos!Liberado = GUsuario
                RBuscaEncabezadoIngresos.Update
                
            MousePointer = 0
            MsgBox "Ingreso Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        DataIngresos.DatabaseName = BasedeDatos
End Sub

Private Sub TxtDoc_Change()
           If IsNumeric(TxtDoc.Text) Then
            DataIngresos.RecordSource = "Select DI.Codigo, I.Descripcion, I.Existencia, DI.Cantidad From DetalleIngresosMateriaPrima as DI, InventarioMateriaPrima as I Where DI.Codigo = I.CodigoMateriaPrima And DI.Bodega = I.Bodega And DI.Documento = " & TxtDoc.Text & " Order By DI.Codigo"
            DataIngresos.Refresh
            DBGridIngresos.Refresh
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE EGRESOS
            Set RBuscaEncabezadoIngresos = Db.OpenRecordset("Select Fecha, Requerido, Observaciones From EncabezadoIngresosMateriaPrima Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoIngresos.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoIngresos!Fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoIngresos!Fecha
                End If
                'REQUERIDO POR
                If IsNull(RBuscaEncabezadoIngresos!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = RBuscaEncabezadoIngresos!Requerido
                End If
                'OBSERVACIONES
                If IsNull(RBuscaEncabezadoIngresos!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RBuscaEncabezadoIngresos!Observaciones
                End If
            Else
                MskFec.Text = ""
                TxtReq.Text = ""
                TxtObs.Text = ""
            End If
      Else
                DataIngresos.RecordSource = "Select DI.Codigo, I.Descripcion, I.Existencia, DI.Cantidad From DetalleIngresosMateriaPrima as DI, InventarioMateriaPrima as I Where DI.Codigo = I.CodigoMateriaPrima And DI.Bodega = I.Bodega And DI.Documento = " & TxtDoc.Text & " Order By DI.Codigo"
                DataIngresos.Refresh
                DBGridIngresos.Refresh
                AnchoColumnas
            
                MskFec.Text = ""
                TxtReq.Text = ""
                TxtObs.Text = ""
      End If


End Sub

Private Sub TxtDoc_GotFocus()
        TxtDoc.SelStart = 0
        TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Sub AnchoColumnas()
        DBGridIngresos.Columns(0).Width = "1100"
        DBGridIngresos.Columns(1).Width = "4000"
        DBGridIngresos.Columns(2).Width = "1100"
        DBGridIngresos.Columns(3).Width = "1000"
End Sub

