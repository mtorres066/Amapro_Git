VERSION 5.00
Begin VB.Form CambiarDocumentosEntradasMP 
   Caption         =   "cambiar documentos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ejecutar 
      Caption         =   "ejecutar"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "CambiarDocumentosEntradasMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REntradas As Recordset
Dim cont As Integer


Private Sub ejecutar_Click()
MousePointer = 1
    Set REntradas = Db.OpenRecordset("Select * From EncabezadoDespachosProductoTerminado Order By Fecha")
        cont = 1
        Do Until REntradas.EOF
                REntradas.Edit
                    REntradas!Documento = cont
                REntradas.Update
                cont = cont + 1
            REntradas.MoveNext
        Loop
MousePointer = 0
        MsgBox "listo"
End Sub
