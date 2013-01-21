VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CambiosUbicacionMateriaPrima 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios Ubicacion De Materia Prima"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "CambiosUbicacionMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataCambiosUbicacion 
      Caption         =   "Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGridCambiosUbicacion 
      Bindings        =   "CambiosUbicacionMateriaPrima.frx":08CA
      Height          =   7095
      Left            =   120
      OleObjectBlob   =   "CambiosUbicacionMateriaPrima.frx":08ED
      TabIndex        =   3
      ToolTipText     =   "escriba la ubicacion de la tarima"
      Top             =   840
      Width           =   11655
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11040
      Picture         =   "CambiosUbicacionMateriaPrima.frx":204B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "salida"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   495
      Left            =   10080
      Picture         =   "CambiosUbicacionMateriaPrima.frx":40BD
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "buscar batch"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Bodega Actual : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   780
      TabIndex        =   7
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label LblBodega 
      BackColor       =   &H00FF8080&
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
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label LblCodigo 
      BackColor       =   &H00FF8080&
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
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "CambiosUbicacionMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaCodigo As Recordset
Dim RBuscaBodega As Recordset

Private Sub CmdConsultar_Click()
        DataCambiosUbicacion.RecordSource = "Select * From DetalleEntradasMateriaPrima Where Codigo = '" & TxtTexto.Text & "' And SaldoDisponibilidad > 0 Order by NumeroIngreso"
        DataCambiosUbicacion.Refresh
        DBGridCambiosUbicacion.Refresh
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridCambiosUbicacion_BeforeUpdate(Cancel As Integer)
On Error Resume Next
        If Err.Number <> 0 Then
            MsgBox Err.Number & " " & Err.Description
        End If
        DataCambiosUbicacion.Recordset!Usuario = GUsuario
End Sub

Private Sub DBGridCambiosUbicacion_KeyDown(KeyCode As Integer, Shift As Integer)
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & DBGridCambiosUbicacion.Columns(2).Text & "'")
            If RBuscaBodega.RecordCount > 0 Then
                LblBodega.Caption = RBuscaBodega!Descripcion
            Else
                LblBodega.Caption = ""
            End If
        
End Sub

Private Sub DBGridCambiosUbicacion_KeyUp(KeyCode As Integer, Shift As Integer)
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & DBGridCambiosUbicacion.Columns(2).Text & "'")
            If RBuscaBodega.RecordCount > 0 Then
                LblBodega.Caption = RBuscaBodega!Descripcion
            Else
                LblBodega.Caption = ""
            End If

End Sub

Private Sub Form_Load()
        DataCambiosUbicacion.Connect = GConnect
        DataCambiosUbicacion.DatabaseName = BasedeDatos
End Sub


Private Sub TxtTexto_Change()
            'BUSCA CODIOG
            Set RBuscaCodigo = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Text & "'")
                If RBuscaCodigo.RecordCount > 0 Then
                    LblCodigo.Caption = RBuscaCodigo!Descripcion
                Else
                    LblCodigo.Caption = ""
                End If

End Sub

Private Sub TxtTexto_GotFocus()
        TxtTexto.SelStart = 0
        TxtTexto.SelLength = Len(TxtTexto.Text)
End Sub

Private Sub TxtTexto_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

