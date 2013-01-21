VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form GeneraInventario 
   Caption         =   "Genera Inventario"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "GeneraInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   17039363
      CurrentDate     =   37938
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   735
      Left            =   4680
      Picture         =   "GeneraInventario.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Generar"
      Height          =   735
      Left            =   3120
      Picture         =   "GeneraInventario.frx":3D6C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   450
   End
   Begin VB.Label LblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Ficha Tecnia"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   930
   End
   Begin VB.Label LblDes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "GeneraInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaInventarioPT As New ADODB.Recordset



Private Sub CmdGenerar_Click()
On Error Resume Next
MousePointer = 11
            'BORRA LOS DATOS DE INVENTARIO QUE YA EXISTAN EN LA FECHA
            'If GOrigenDeDatos = "AmaproAccess" Then
            '    Conexion.Execute "Delete from Inventario where Fecha = #" & Format(DtpFecha.Value, "mm/dd/yyyy") & "#"
            'Else
            '    Conexion.Execute "Delete from Inventario where Fecha = To_Date('" & DtpFecha.Value & "', 'dd/mm/yyyy')"
            'End If
            'SI HAY ERROR
            'If Err <> 0 Then
            '    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            'End If
        
            'BUSCA EL INVENTARIO
            Set RBuscaInventarioPT = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaInventarioPT, "Select D.FichaTecnica, D.Bodega, Sum(D.Saldo) From DetalleEntradasInventario D, FichaTecnica F Where D.FichaTecnica = F.Esp_Tec And F.TipoInventario = 'PRODUCTO TERMINADO' And D.FichaTecnica Like '" & TxtCod.Text & "%' And D.Saldo > 0 And (D.Bodega = 'T02' OR D.Bodega = '001') Group By D.FichaTecnica, D.Bodega")
                Else
                    Call Abrir_Recordset(RBuscaInventarioPT, "Select D.FichaTecnica, D.Bodega, Sum(D.Saldo) From DetalleEntradasInventario D, FichaTecnica F Where UPPER(D.FichaTecnica) = UPPER(F.Esp_tec) And UPPER(F.TipoInventario) = 'PRODUCTO TERMINADO' And UPPER(D.FichaTecnica) Like '" & UCase(TxtCod.Text) & "%' And D.Saldo > 0 And Mid(D.Bodega ,1,1) = 'T' OR D.Bodega = '001' Group By D.FichaTecnica, D.Bodega")
                End If
                If RBuscaInventarioPT.RecordCount > 0 Then
                    'INICIA LA TRANSACCION
                    Conexion.BeginTrans
                            Do Until RBuscaInventarioPT.EOF
                                    'INSERTA LA INFORMACION
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Conexion.Execute "Insert Into Inventario (Fecha, FichaTecnica, Bodega, Cantidad, Usuario) VALUES(#" & Format(DtpFecha.Value, "mm/dd/yyyy") & "#, '" & RBuscaInventarioPT(0) & "', '" & RBuscaInventarioPT(1) & "', " & RBuscaInventarioPT(2) & ", '" & GUsuario & "')"
                                    Else
                                        Conexion.Execute "Insert Into Inventario (Fecha, FichaTecnica, Bodega, Cantidad, Usuario) VALUES(To_Date('" & DtpFecha.Value & "', 'dd/mm/yyyy')" & ", '" & RBuscaInventarioPT(0) & "', '" & RBuscaInventarioPT(1) & "', " & RBuscaInventarioPT(2) & ", '" & GUsuario & "')"
                                    End If
                                    'SI HAY ERROR
                                    If Err <> 0 Then
                                        'Conexion.RollbackTrans
                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                        Err.Clear
                                        MousePointer = 0
                                        'Exit Sub
                                    End If
                                RBuscaInventarioPT.MoveNext
                            Loop
                    'TERMINA LA TRANSACCION
                    Conexion.CommitTrans
                Else
                End If
            
MousePointer = 0
        
        MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        DtpFecha.Value = Date
End Sub

Private Sub TxtCod_Change()
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCod.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCod.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblDes.Caption = ""
                End If
        
End Sub

Private Sub TxtCod_GotFocus()
        TxtCod.SelStart = 0
        TxtCod.SelLength = Len(TxtCod.Text)
End Sub
