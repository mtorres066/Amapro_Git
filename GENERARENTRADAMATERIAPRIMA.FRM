VERSION 5.00
Begin VB.Form GenerarEntradaMateriaPrima 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Entrada De Materia Prima"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "GenerarEntradaMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salida"
      Height          =   1095
      Left            =   5160
      Picture         =   "GenerarEntradaMateriaPrima.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Generar"
      Default         =   -1  'True
      Height          =   1095
      Left            =   5160
      Picture         =   "GenerarEntradaMateriaPrima.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox TxtDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento De Entrada"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5340
   End
   Begin VB.Image Image1 
      Height          =   3765
      Left            =   360
      Picture         =   "GenerarEntradaMateriaPrima.frx":0A56
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "GenerarEntradaMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaDetalleEntradasMateriaPrima As Recordset
Dim RSupervisionMateriaPrima As Recordset
Dim RBuscaSupervisionMateriaPrima As Recordset


Private Sub CmdGenerar_Click()
On Error Resume Next
    MousePointer = 11
            
            'BUSCA SI YA EXISTE EL DOCUMENTO EN MATERIA PRIMA
            Set RBuscaSupervisionMateriaPrima = Db.OpenRecordset("Select * from SupervisionEntradasMateriaPrima Where Documento = '" & TxtDoc.Text & "'")
            If RBuscaSupervisionMateriaPrima.RecordCount > 0 Then
                    MsgBox "Documento De Entrada Ya Esta Generado", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
            End If
            
            'INICIALIZA EL RECORDSET PARA LA BASE DE DATOS
            Set RSupervisionMateriaPrima = Db.OpenRecordset("SupervisionEntradasMateriaPrima")
            
            'BUSCA EL DETALLE DE LA ENTRADA DE LA MATERIA PRIMA DE ACUERDO AL DOCUMENTO REQUERIDO
            Set RBuscaDetalleEntradasMateriaPrima = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Documento = '" & TxtDoc.Text & "'")
                If RBuscaDetalleEntradasMateriaPrima.RecordCount > 0 Then
                             
                Do Until RBuscaDetalleEntradasMateriaPrima.EOF
                    'AGREGA DATOS
                    RSupervisionMateriaPrima.AddNew
                            RSupervisionMateriaPrima!Documento = RBuscaDetalleEntradasMateriaPrima!Documento
                            RSupervisionMateriaPrima!Cantidad = RBuscaDetalleEntradasMateriaPrima!Cantidad
                            RSupervisionMateriaPrima!Bodega = RBuscaDetalleEntradasMateriaPrima!Bodega
                            RSupervisionMateriaPrima!Codigo = RBuscaDetalleEntradasMateriaPrima!Codigo
                            RSupervisionMateriaPrima!DocumentoPedido = RBuscaDetalleEntradasMateriaPrima!DocumentoPedido
                            RSupervisionMateriaPrima!Bulto = RBuscaDetalleEntradasMateriaPrima!Bulto
                            'LA FECHA, EL CAMPO CALIDAD, ADVERTENCIA1, ADVERTENCIA2, ADVERTENCIA3, NOCEDULA, ESTADO
                            'LOS ASIGNA EL DE CONTROL DE CALIDAD A LA HORA DE SUPERVISAR
                    RSupervisionMateriaPrima.Update
                    
                    If Err <> 0 Then
                    End If
                    
                    RBuscaDetalleEntradasMateriaPrima.MoveNext
                Loop
            
            Else
                MsgBox "Numero De Documento No Existe En Entradas De Materia Prima", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
                Exit Sub
            End If
            
                MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
    MousePointer = 0
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
