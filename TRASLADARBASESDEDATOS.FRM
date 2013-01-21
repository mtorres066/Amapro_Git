VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form TrasladarBasesDeDatos 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "ok"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   3960
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProduccionLiberadaConTarimas"
      Top             =   6000
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TRASLADARBASESDEDATOS.frx":0000
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "TRASLADARBASESDEDATOS.frx":0014
      TabIndex        =   1
      Top             =   4200
      Width           =   9735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProduccionTotal"
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TRASLADARBASESDEDATOS.frx":09EF
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "TRASLADARBASESDEDATOS.frx":0A03
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "TrasladarBasesDeDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
On Error Resume Next
        Do Until Data1.Recordset.EOF
                Data2.Recordset.AddNew
                        Data2.Recordset!Fec_prd = Data1.Recordset!Fec_prd
                        Data2.Recordset!Linea = Data1.Recordset!Linea
                        Data2.Recordset!Esp_tec = Data1.Recordset!Esp_tec
                        Data2.Recordset!Tarima = Data1.Recordset!Tarima
                        Data2.Recordset!Fec_prdL = Data1.Recordset!FechaTarCom
                        Data2.Recordset!LineaL = Data1.Recordset!LineaCom
                        Data2.Recordset!Esp_tecL = Data1.Recordset!FichaTecnicaCom
                        Data2.Recordset!TarimaL = Data1.Recordset!TarimaCom
                        Data2.Recordset!CalidadL = "C"
                        Data2.Recordset!Revisados = Data1.Recordset!EnvCom
                        Data2.Recordset!NoConforme = 0
                        Data2.Recordset!Liberados = Data1.Recordset!EnvCom0
                        Data2.Recordset!EnTarima = Data1.Recordset!EnvCom0
                        Data2.Recordset!Minutos = 0
                        
                        'Data2.Recordset(10) = Data1.Recordset(10)
                        'Data2.Recordset(11) = Data1.Recordset(11)
                        'Data2.Recordset(12) = Data1.Recordset(12)
                        'Data2.Recordset(13) = Data1.Recordset(13)
                        'Data2.Recordset(14) = Data1.Recordset(14)
                        'Data2.Recordset(15) = Data1.Recordset(15)
                        'Data2.Recordset(16) = Data1.Recordset(16)
                        'Data2.Recordset(17) = Data1.Recordset(17)
                        'Data2.Recordset(18) = Data1.Recordset(18)
                        'Data2.Recordset(19) = Data1.Recordset(19)
                        'Data2.Recordset(20) = Data1.Recordset(20)
                        'Data2.Recordset(21) = Data1.Recordset(21)
                        'Data2.Recordset(22) = Data1.Recordset(22)
                        'Data2.Recordset(23) = Data1.Recordset(23)
                        'Data2.Recordset(24) = Data1.Recordset(24)
                        'Data2.Recordset(25) = Data1.Recordset(25)
                        'Data2.Recordset(26) = Data1.Recordset(26)
                        'Data2.Recordset(27) = Data1.Recordset(27)
                        'Data2.Recordset(28) = Data1.Recordset(28)
                        'Data2.Recordset(29) = Data1.Recordset(29)
                        'Data2.Recordset(30) = Data1.Recordset(30)
                        'Data2.Recordset(31) = Data1.Recordset(31)
                        'Data2.Recordset(32) = Data1.Recordset(32)
                        'Data2.Recordset(33) = Data1.Recordset(33)
                        'Data2.Recordset(34) = Data1.Recordset(34)
                        'Data2.Recordset(35) = Data1.Recordset(35)
                        'Data2.Recordset(36) = Data1.Recordset(36)
                        'Data2.Recordset(37) = Data1.Recordset(37)
                Data2.Recordset.Update
                
                        If Err <> 0 Then
                            
                        End If
            Data1.Recordset.MoveNext
        Loop
        
        MsgBox "listo"
End Sub
