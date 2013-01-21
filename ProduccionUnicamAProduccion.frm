VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ProduccionUnicamaProduccion 
   Caption         =   "Cambiar Produccion Unicam a Produccion"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Batch Unicam a Batch"
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   5160
      Width           =   5295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pasar De Ficha Tecnica a Ficha Tecnica Con Materia Prima"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   7320
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pasa Materias Primas Y Numero Ingreso a ProduccionConMateriaPrima"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   6600
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pasa Defectos a ProduccionConDefectos"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6000
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "pasa unicam o encajado a produccion"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BatchDatos"
      Top             =   8040
      Width           =   8895
   End
   Begin VB.Data DataUnicam 
      Caption         =   "Unicam O Encajado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BatchDatosUnicam"
      Top             =   3960
      Width           =   3975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ProduccionUnicamAProduccion.frx":0000
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "ProduccionUnicamAProduccion.frx":0019
      TabIndex        =   4
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "ProduccionUnicamaProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RProduccion As Recordset
Dim RProduccionDefectos As Recordset
Dim RProduccionMateriasPrimas As Recordset

Dim RFichaTecnica As Recordset
Dim RFichaTecnicaConMateriaPrima As Recordset



Private Sub Command1_Click()
MousePointer = 11
        Do Until DataUnicam.EOFAction
                
                DataProduccion.Recordset.AddNew
                        'PRODUCCION UNICAM PARA PRODUCCION
                        'DataProduccion.Recordset!Fec_prd = DataUnicam.Recordset!Fec_prd
                        'DataProduccion.Recordset!Hor_prd = DataUnicam.Recordset!Hor_prd
                        'DataProduccion.Recordset!Linea = DataUnicam.Recordset!Linea
                        'DataProduccion.Recordset!Tarima = DataUnicam.Recordset!Tarima
                        'DataProduccion.Recordset!Esp_Tec = DataUnicam.Recordset!Esp_Tec
                        'DataProduccion.Recordset!Batch = DataUnicam.Recordset!Batch
                        'DataProduccion.Recordset!Envases = DataUnicam.Recordset!Envases
                        'DataProduccion.Recordset!Calidad = DataUnicam.Recordset!Calidad
                        'DataProduccion.Recordset!Muestra = DataUnicam.Recordset!Muestra
                        'DataProduccion.Recordset!Turno = DataUnicam.Recordset!Turno
                        'DataProduccion.Recordset!Cod_Emp = DataUnicam.Recordset!Cod_Emp
                        'DataProduccion.Recordset!NoMP9301 = DataUnicam.Recordset!NoMP9301
                        'DataProduccion.Recordset!ColorMP9301 = DataUnicam.Recordset!ColorMP9301
                        'DataProduccion.Recordset!Observaciones = DataUnicam.Recordset!Observaciones
                        
                        'PRODUCCION ENCAJADA PARA PRODUCCION
                        'DataProduccion.Recordset!Fec_prd = DataUnicam.Recordset!Fec_prd
                        'DataProduccion.Recordset!Hor_prd = DataUnicam.Recordset!Hor_prd
                        'DataProduccion.Recordset!Linea = DataUnicam.Recordset!Linea
                        'DataProduccion.Recordset!Tarima = DataUnicam.Recordset!Tarima
                        'DataProduccion.Recordset!Esp_Tec = DataUnicam.Recordset!Esp_Tec
                        'DataProduccion.Recordset!Batch = DataUnicam.Recordset!Batch
                        'DataProduccion.Recordset!Envases = DataUnicam.Recordset!Envases
                        'DataProduccion.Recordset!Calidad = DataUnicam.Recordset!Calidad
                        'DataProduccion.Recordset!Turno = DataUnicam.Recordset!Turno
                        'DataProduccion.Recordset!Cod_Emp = DataUnicam.Recordset!Usuario
                        'DataProduccion.Recordset!NoMP9301 = DataUnicam.Recordset!NoMP9301
                        'DataProduccion.Recordset!ColorMP9301 = DataUnicam.Recordset!ColorMP9301
                        'DataProduccion.Recordset!Observaciones = DataUnicam.Recordset!Observaciones
                        
                        'BATCH UNICAM PARA BATCH
                        'DataProduccion.Recordset(0) = DataUnicam.Recordset(0)
                        'DataProduccion.Recordset(1) = DataUnicam.Recordset(1)
                        'DataProduccion.Recordset(2) = DataUnicam.Recordset(2)
                        'DataProduccion.Recordset(3) = DataUnicam.Recordset(3)
                        'DataProduccion.Recordset(4) = DataUnicam.Recordset(4)
                        'DataProduccion.Recordset(5) = DataUnicam.Recordset(5)
                        'DataProduccion.Recordset(6) = DataUnicam.Recordset(6)
                        'DataProduccion.Recordset(7) = DataUnicam.Recordset(7)
                        
                        
                        'BATCHDATOSUNICAM PARA BATCHDATOS
                        'DataProduccion.Recordset(0) = DataUnicam.Recordset(0)
                        'DataProduccion.Recordset(1) = DataUnicam.Recordset(1)
                        'DataProduccion.Recordset(2) = DataUnicam.Recordset(2)
                        'DataProduccion.Recordset(3) = DataUnicam.Recordset(3)
                        'DataProduccion.Recordset(4) = DataUnicam.Recordset(4)
                        'DataProduccion.Recordset(5) = DataUnicam.Recordset(5)
                        'DataProduccion.Recordset(6) = DataUnicam.Recordset(6)
                        'DataProduccion.Recordset(7) = DataUnicam.Recordset(7)
                        'DataProduccion.Recordset(8) = DataUnicam.Recordset(8)
                        'DataProduccion.Recordset(9) = DataUnicam.Recordset(9)
                        'DataProduccion.Recordset(10) = DataUnicam.Recordset(10)
                        'DataProduccion.Recordset(11) = DataUnicam.Recordset(11)
                        'DataProduccion.Recordset(12) = DataUnicam.Recordset(12)
                        
                        
                DataProduccion.Recordset.Update
                
            DataUnicam.Recordset.MoveNext
        Loop
        
MousePointer = 0
                MsgBox "ya"

End Sub

Private Sub Command2_Click()
On Error Resume Next
            Set RProduccionDefectos = Db.OpenRecordset("Select * From ProduccionConDefectos")
            
            Set RProduccion = Db.OpenRecordset("Select * From Produccion")
            
            Do Until RProduccion.EOF
                    If RProduccion!Defecto3 = "" Or IsNull(RProduccion!Defecto3) Then
                    Else
                        RProduccionDefectos.AddNew
                            RProduccionDefectos!Fec_prd = RProduccion!Fec_prd
                            RProduccionDefectos!Linea = RProduccion!Linea
                            RProduccionDefectos!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionDefectos!Tarima = RProduccion!Tarima
                            RProduccionDefectos!Defecto = RProduccion!Defecto3
                            RProduccionDefectos!Cantidad = RProduccion!Cantidad3
                        RProduccionDefectos.Update
                        If Err <> 0 Then
                            MsgBox Err.Number & Err.Description
                        End If
                    End If
                RProduccion.MoveNext
            Loop
            MsgBox "Ya"
End Sub

Private Sub Command3_Click()
On Error Resume Next
            Set RProduccionMateriasPrimas = Db.OpenRecordset("Select * From ProduccionConMateriaPrima")
            
            Set RProduccion = Db.OpenRecordset("Select * From Produccion")
            
            Do Until RProduccion.EOF
                    'FONDO
                    If RProduccion!Fondo = "" Or IsNull(RProduccion!Fondo) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!Fondo
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoFondo
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    'HOJALATA
                    If RProduccion!Platina = "" Or IsNull(RProduccion!Platina) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!Platina
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoHojalata
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'ALAMBRE DE COBRE
                    If RProduccion!AlambreCobre = "" Or IsNull(RProduccion!AlambreCobre) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!AlambreCobre
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoAlambre
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'BARNIZ LIQUIDO
                    If RProduccion!BarnizLiquido = "" Or IsNull(RProduccion!BarnizLiquido) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!BarnizLiquido
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoBarnizL
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'BARNIZ POLVO
                    If RProduccion!BarnizPolvo = "" Or IsNull(RProduccion!BarnizPolvo) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!BarnizPolvo
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoBarnizP
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'SELLO SOLVENTE
                    If RProduccion!SelloSolve = "" Or IsNull(RProduccion!SelloSolve) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!SelloSolve
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoSelloSOlve
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'NYLON
                    If RProduccion!NylonStrech = "" Or IsNull(RProduccion!NylonStrech) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!NylonStrech
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoNylon
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                    'BOBINA
                    If RProduccion!CodigoBobina = "" Or IsNull(RProduccion!CodigoBobina) Then
                    Else
                        RProduccionMateriasPrimas.AddNew
                            RProduccionMateriasPrimas!Fec_prd = RProduccion!Fec_prd
                            RProduccionMateriasPrimas!Linea = RProduccion!Linea
                            RProduccionMateriasPrimas!Esp_Tec = RProduccion!Esp_Tec
                            RProduccionMateriasPrimas!Tarima = RProduccion!Tarima
                            RProduccionMateriasPrimas!CodigoMateriaPrima = RProduccion!CodigoBobina
                            RProduccionMateriasPrimas!Bulto = RProduccion!CodigoIngresoBobina
                        RProduccionMateriasPrimas.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                    End If
                    
                RProduccion.MoveNext
            Loop
            MsgBox "Ya"

End Sub

Private Sub Command4_Click()
On Error Resume Next
            Set RFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica")
            
            Set RFichaTecnicaConMateriaPrima = Db.OpenRecordset("Select * From FichaTecnicaConMateriaPrima")
            
            
            Do Until RFichaTecnica.EOF
                        'PLATINA
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!Platina
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'FONDO
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!Fondo
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'ALAMBRE
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodAla
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'BARNIZ LIQUIDO
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodBarLiq
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'BARNIZ POLVO
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodBarPol
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'SELLO SOLVENTE
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodselSol
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'NYLON STRECH
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodNylStr
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                        'BOBINA
                        RFichaTecnicaConMateriaPrima.AddNew
                            RFichaTecnicaConMateriaPrima!Esp_Tec = RFichaTecnica!Esp_Tec
                            RFichaTecnicaConMateriaPrima!CodigoMateriaPrima = RFichaTecnica!CodBobina
                        RFichaTecnicaConMateriaPrima.Update
                        If Err <> 0 Then
                            'MsgBox Err.Number & Err.Description
                        End If
                
                RFichaTecnica.MoveNext
            Loop
            
            MsgBox "ya"
            
            
            
End Sub

Private Sub Command5_Click()
MousePointer = 11
        Do Until DataUnicam.EOFAction
                
                DataProduccion.Recordset.AddNew
                        'BATCH UNICAM PARA BATCH
                        'DataProduccion.Recordset(0) = DataUnicam.Recordset(0)
                        'DataProduccion.Recordset(1) = DataUnicam.Recordset(1)
                        'DataProduccion.Recordset(2) = DataUnicam.Recordset(2)
                        'DataProduccion.Recordset(3) = DataUnicam.Recordset(3)
                        'DataProduccion.Recordset(4) = DataUnicam.Recordset(4)
                        'DataProduccion.Recordset(5) = DataUnicam.Recordset(5)
                        'DataProduccion.Recordset(6) = DataUnicam.Recordset(6)
                        'DataProduccion.Recordset(7) = DataUnicam.Recordset(7)
                        
                        
                        'BATCHDATOSUNICAM PARA BATCHDATOS
                        DataProduccion.Recordset(0) = DataUnicam.Recordset(0)
                        DataProduccion.Recordset(1) = DataUnicam.Recordset(1)
                        DataProduccion.Recordset(2) = DataUnicam.Recordset(2)
                        DataProduccion.Recordset(3) = DataUnicam.Recordset(3)
                        DataProduccion.Recordset(4) = DataUnicam.Recordset(4)
                        DataProduccion.Recordset(5) = DataUnicam.Recordset(5)
                        DataProduccion.Recordset(6) = DataUnicam.Recordset(6)
                        DataProduccion.Recordset(7) = DataUnicam.Recordset(7)
                        DataProduccion.Recordset(8) = DataUnicam.Recordset(8)
                        DataProduccion.Recordset(9) = DataUnicam.Recordset(9)
                        DataProduccion.Recordset(10) = DataUnicam.Recordset(10)
                        DataProduccion.Recordset(11) = DataUnicam.Recordset(11)
                        DataProduccion.Recordset(12) = DataUnicam.Recordset(12)
                                                
                DataProduccion.Recordset.Update
                
            DataUnicam.Recordset.MoveNext
        Loop
        
MousePointer = 0
                MsgBox "ya"

End Sub
