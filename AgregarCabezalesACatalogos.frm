VERSION 5.00
Begin VB.Form AgregarCabezalesACatalogos 
   Caption         =   "Agregar Cabezales A Catalogos"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   975
      Left            =   1440
      Picture         =   "AgregarCabezalesACatalogos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "AgregarCabezalesACatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RRutinas As New ADODB.Recordset


Private Sub Command1_Click()

MousePointer = 11
        Set RRutinas = New ADODB.Recordset
            Call Abrir_Recordset(RRutinas, "select * From Rutinas")
                If RRutinas.RecordCount > 0 Then
                    Do Until RRutinas.EOF
                            Conexion.Execute "update variablesmedia set cabezales = " & RRutinas!cabezal & " Where Rutina = '" & RRutinas!Rutina & "'"
                                If Err <> 0 Then
                                    MousePointer = 0
                                    MsgBox "Error " & Err.Description
                                    Exit Sub
                                End If
                            
                        
                        RRutinas.MoveNext
                    Loop
                End If
MousePointer = 0
            MsgBox "Datos Actualizados Con Exito", vbOKOnly + vbInformation, "Informacion"

End Sub
