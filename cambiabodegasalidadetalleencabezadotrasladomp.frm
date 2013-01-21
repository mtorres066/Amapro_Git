VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rencabezado As Recordset
Dim rdetalle As Recordset
Private Sub Command1_Click()

        Set rencabezado = Db.OpenRecordset("Select documento, BodegaSalida From EncabezadoTrasladosMateriaPrimap")
            If rencabezado.RecordCount > 0 Then
                Do Until rencabezado.EOF
                        Set rdetalle = Db.OpenRecordset("Select documento, BodegaSalida From DetalleTrasladosMateriaPrimaP Where Documento = " & rencabezado!Documento)
                            If rdetalle.RecordCount > 0 Then
                                rencabezado.Edit
                                    rencabezado!BodegaSalida = rdetalle!BodegaSalida
                                rencabezado.Update
                            End If
                    rencabezado.MoveNext
                Loop
            End If
MsgBox "ya"
End Sub
