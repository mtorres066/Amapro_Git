VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form EmpleadosFactores 
   Caption         =   "Factores Para Calculos De Horas Extras"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataFactores 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EmpleadosFactores"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGridFactores 
      Bindings        =   "EmpleadosFactores.frx":0000
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "EmpleadosFactores.frx":001B
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "EmpleadosFactores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGridFactores_BeforeUpdate(Cancel As Integer)
On Error Resume Next
        If Err.Number <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description
        End If
End Sub

Private Sub Form_Load()
    DataFactores.ConnectionString = GTipoProveedor
    DataFactores.Refresh

End Sub
