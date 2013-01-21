VERSION 5.00
Begin VB.Form ActualizaParos 
   Caption         =   "Actualiza Paros"
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
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "ActualizaParos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RDatos As New ADODB.Recordset
Private Sub Command1_Click()
MousePointer = 11
            
            
            Conexion.Execute "Update encabezadocapturaparos set ParoCF = 0, ParoMP = 0"
                            If Err <> 0 Then
                                MsgBox "Error En Actualizar a Cero" & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                                Exit Sub
                            End If
            

            Set RDatos = New ADODB.Recordset
                Call Abrir_Recordset(RDatos, "Select E.Documento, sum(DC.minutos) from DetalleCapturaParos DC, EncabezadoCapturaParos E, Paros P where E.Documento = DC.Documento And DC.Paro = P.CodigoParo And P.Tipo2 = 'CF' Group By E.Documento")
                    If RDatos.RecordCount > 0 Then
                                Do Until RDatos.EOF
                                                Conexion.Execute "Update encabezadocapturaparos set ParoCF = " & RDatos(1) & " Where Documento = " & RDatos(0)
                                                If Err <> 0 Then
                                                    MsgBox "Error En Actualizar CF" & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                    Err.Clear
                                                    Exit Sub
                                                End If
                                
                                        RDatos.MoveNext
                                Loop
                    End If
                    
            
            Set RDatos = New ADODB.Recordset
                Call Abrir_Recordset(RDatos, "Select E.Documento, sum(DC.minutos) from DetalleCapturaParos DC, EncabezadoCapturaParos E, Paros P where E.Documento = DC.Documento And DC.Paro = P.CodigoParo And P.Tipo2 = 'MP' Group By E.Documento")
                    If RDatos.RecordCount > 0 Then
                                Do Until RDatos.EOF
                                                Conexion.Execute "Update encabezadocapturaparos set ParoMP = " & RDatos(1) & " Where Documento = " & RDatos(0)
                                                If Err <> 0 Then
                                                    MsgBox "Error En Actualizar MP " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                    Err.Clear
                                                    Exit Sub
                                                End If
                                
                                        RDatos.MoveNext
                                Loop
                    End If
                    
                    MsgBox "ya"
                    
    MousePointer = 0
End Sub
