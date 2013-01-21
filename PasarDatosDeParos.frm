VERSION 5.00
Begin VB.Form PasarDatosDeParos 
   BackColor       =   &H000000FF&
   Caption         =   "Pasar Datos De Paros"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "PasarDatosDeParos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RDetalle As New ADODB.Recordset


Private Sub Command1_Click()
On Error Resume Next
MousePointer = 11
        Set RDetalle = New ADODB.Recordset
            Call Abrir_Recordset(RDetalle, "Select * From DetalleSalidasInventario Where Documento = 3579")
                    If RDetalle.RecordCount > 0 Then
                        Conexion.BeginTrans
                        Do Until RDetalle.EOF
                                Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo + " & RDetalle!Cantidad & " Where FechaProduccion = #" & Format(RDetalle!FechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & RDetalle!Linea & "' And FichaTecnica = '" & RDetalle!FichaTecnica & "' And Tarima = " & RDetalle!Tarima
                                                
                                        If Err <> 0 Then
                                               MousePointer = 0
                                               MsgBox "No Se Grabaran Los Cambios " & Err.Description
                                               Conexion.RollbackTrans
                                               Err.Clear
                                        End If
                
                                                
                        
                            RDetalle.MoveNext
                        Loop
                        
                        Conexion.CommitTrans
                        MousePointer = 0
                        MsgBox "No se como pero se salvaron señor miguel"
                    End If

End Sub

Private Sub Form_Load()
            GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=\\serverculiacan\Amapro\Metalenvases.mdb; Jet OLEDB:Database Password=metal"
            'GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & App.Path & "\Metalenvases.mdb; Jet OLEDB:Database Password=metal"
                                
            
            'INICIALIZA O CREA LA INSTANCIA DE LA CONECCION
            Set Conexion = New ADODB.Connection
            Conexion.ConnectionString = GConectionString
            Conexion.Open
End Sub
