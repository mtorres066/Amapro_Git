VERSION 5.00
Begin VB.Form CambiarClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Password"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   Icon            =   "CambiarClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNuevo 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   615
      Left            =   2400
      Picture         =   "CambiarClave.frx":2072
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Cambiar "
      Height          =   615
      Left            =   1560
      Picture         =   "CambiarClave.frx":40E4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Txtviejo 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox TxtAsociado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password Nuevo"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password Anterior"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "CambiarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cont As Integer
Dim VLargo As Integer
Dim VNumeroAscii As String
Dim VEncriptado As String
Dim RBuscaAsociado As New ADODB.Recordset

Private Sub CmdCambiar_Click()

MousePointer = 11
                    '----------------------------------------------------------------------------------------------------
                    'PROCESO PARA ENCRIPTAR EL PASSWORD AGARRAMOS CADA LETRA DEL PASSWORD Y LE ASIGNAMOS EL CODIGO ASSCII
                    'Y TAMBIEN LE AGREGAMOS UN NUMERO CUALQUIERA (0110) PARA QUE SEA UN POCO MAS DIFICIL DE LEERLO
                    
                    Cont = 1
                    VLargo = Len(Txtviejo.Text)
                    VNumeroAscii = ""
                    VEncriptado = ""
                    
                    Do While Cont <= VLargo
                       VNumeroAscii = Asc(Mid(Txtviejo.Text, Cont, 1))
                       VEncriptado = VEncriptado & VNumeroAscii & "0110"
                       Cont = Cont + 1
                    Loop
        
       'VERIFICA EL NOMBRE DEL ASOSCIADO
       Set RBuscaAsociado = New ADODB.Recordset
       Call Abrir_Recordset(RBuscaAsociado, "Select * From Usuarios Where Usuario = '" & TxtAsociado.Text & "'")
        If RBuscaAsociado.RecordCount > 0 Then
        Else
            MsgBox "Usuario No Existe"
            TxtAsociado.SetFocus
            Exit Sub
        End If
        
        'VERIFICA EL ASOCIADO CON SU PASSWORD
        Set RBuscaAsociado = New ADODB.Recordset
        Call Abrir_Recordset(RBuscaAsociado, "select * from Usuarios where Usuario = '" & TxtAsociado.Text & "'" & " and Clave = '" & VEncriptado & "'")
            If RBuscaAsociado.RecordCount > 0 Then
            Else
                MsgBox "Password Anterior Incorrecto Para El Usuario", vbOKOnly + vbInformation, "Informacion"
                Txtviejo.SetFocus
                Exit Sub
            End If
        
        
            
      '----------------------------------------------------------------------------------------------------
                    'PROCESO PARA ENCRIPTAR EL PASSWORD AGARRAMOS CADA LETRA DEL PASSWORD Y LE ASIGNAMOS EL CODIGO ASSCII
                    'Y TAMBIEN LE AGREGAMOS UN NUMERO CUALQUIERA (0110) PARA QUE SEA UN POCO MAS DIFICIL DE LEERLO
                    
                    Cont = 1
                    VLargo = Len(TxtNuevo.Text)
                    VNumeroAscii = ""
                    VEncriptado = ""
                    
                    Do While Cont <= VLargo
                       VNumeroAscii = Asc(Mid(TxtNuevo.Text, Cont, 1))
                       VEncriptado = VEncriptado & VNumeroAscii & "0110"
                       Cont = Cont + 1
                    Loop
                                        
                    'ACTUALIZA EL NUEVO PASSWORD
                    Conexion.Execute ("Update Usuarios Set Clave = '" & VEncriptado & "' Where Usuario = '" & TxtAsociado.Text & "'")
            
MousePointer = 0

            MsgBox "Password Cambiado", vbOKOnly + vbInformation, "Informacion"

                    
                                        
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub TxtAsociado_GotFocus()
    TxtAsociado.SelStart = 0
    TxtAsociado.SelLength = Len(TxtAsociado.Text)
End Sub

Private Sub TxtAsociado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TxtNuevo_GotFocus()
        TxtNuevo.SelStart = 0
        TxtNuevo.SelLength = Len(TxtNuevo.Text)
End Sub

Private Sub TxtNuevo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub Txtviejo_GotFocus()
        Txtviejo.SelStart = 0
        Txtviejo.SelLength = Len(Txtviejo.Text)
End Sub

Private Sub Txtviejo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
