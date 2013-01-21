VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form actualizallavecapturarutinas 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar bar 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "actualizar llave "
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "actualizallavecapturarutinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New ADODB.Recordset
Dim Cont As Long

Private Sub Command1_Click()
On Error Resume Next
    
    bar.Value = 1

    Set r = New ADODB.Recordset
        Call Abrir_Recordset(r, "select * from capturarutinas where isnull(llave) order by fec_rut")
            If r.RecordCount > 0 Then
                    bar.Max = r.RecordCount
                    Cont = 1
                    Do Until Cont = 10000
                            r!llave = Cont
                            r.Update
                            If Err <> 0 Then
                                MsgBox "error " & Err.Description
                                Err.Clear
                                Exit Sub
                            End If
                            bar.Value = Cont
                            Cont = Cont + 1
                        r.MoveNext
                    Loop
                    MsgBox "Datos Actualizados Con Exito", vbOKOnly + vbInformation, "informacion"
            Else
                MsgBox "ya no hay datos para actualizar"
                    
            End If

End Sub

