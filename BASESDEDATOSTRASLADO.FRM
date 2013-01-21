VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "cambiar datos"
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   8160
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "BASESDEDATOSTRASLADO.frx":0000
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "BASESDEDATOSTRASLADO.frx":0014
      TabIndex        =   2
      Top             =   3840
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "BASESDEDATOSTRASLADO.frx":09E7
      Height          =   3375
      Left            =   0
      OleObjectBlob   =   "BASESDEDATOSTRASLADO.frx":09FB
      TabIndex        =   1
      Top             =   0
      Width           =   11775
   End
   Begin VB.Data Data2 
      Caption         =   "amapro"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasMateriaPrima"
      Top             =   8040
      Width           =   9015
   End
   Begin VB.Data Data1 
      Caption         =   "backup"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\backup.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasMateriaPrima"
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Recordset


Private Sub Command1_Click()
On Error Resume Next
    Do Until Data1.Recordset.EOF
            
            Set r = Db.OpenRecordset("Select Disponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & Data1.Recordset!Codigo & "' And NumeroIngreso = " & Data1.Recordset!NumeroIngreso)
                If r.RecordCount > 0 Then
                    r.Edit
                        r!Disponibilidad = Data1.Recordset!Disponibilidad
                    r.Update
                End If
            Data1.Recordset.MoveNext
    Loop
            
            
            
            MsgBox "ya"
            
            'Data2.Recordset.AddNew
                'Data2.Recordset!CodigoProducto = Data1.Recordset!Esp_Tec
                'Data2.Recordset!Descripcion = Data1.Recordset!Descrip
                'Data2.Recordset!Bodega = "003"
                'Data2.Recordset!FechaInventario = "23/02/2002"
                'Data2.Recordset!UnidadMedida = "ENVASES"
                'Data2.Recordset!UsuarioAgregar = "ERICK"
                'Data2.Recordset!UsuarioEditar = "ERICK"
                
            '    Data2.Recordset!Documento = "1"
            '    Data2.Recordset!Bodega = "001"
            '    Data2.Recordset!Codigo = Data1.Recordset!CodigoMateriaPrima
            '    Data2.Recordset!Cantidad = Data1.Recordset!Cantidad
            '    Data2.Recordset!DocumentoPedido = "1"
            '    Data2.Recordset!NumeroIngreso = Data1.Recordset!codigoingreso
            '    Data2.Recordset!Calidad = "ACEPTADA"
            '    Data2.Recordset!Estado = "LIBERADO"
            '    Data2.Recordset!Disponibilidad = "EXISTE"
            '    Data2.Recordset!BodegaDisponibilidad = "001"
            '    Data2.Recordset!CantidadTraslado = Data1.Recordset!Cantidad
            '    Data2.Recordset!SaldoDisponibilidad = Data1.Recordset!Cantidad
            '    Data2.Recordset!CantidadSalida = 0
           ' D 'ata2.Recordset.Update
            
            
            '    Set r = Db.OpenRecordset("Select Inicio, Existencia From InventarioMateriaPrima where Codigomateriaprima = '" & Data1.Recordset!CodigoMateriaPrima & "' And Bodega = '" & Data1.Recordset!Bodega & "'")
            '    If r.RecordCount > 0 Then
            '        r.Edit
            '            r!Inicio = Data1.Recordset!Cantidad
            '            r!Existencia = Data1.Recordset!Cantidad
            '        r.Update
            '    End If
                
                If Err <> 0 Then
                    MsgBox Err.Description
                End If
                
        'Data1.Recordset.MoveNext
    'Loop
    
    'MsgBox "ya"
    
End Sub

