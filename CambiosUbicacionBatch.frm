VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CambiosUbicacionBatch 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios De Ubicacion Producto Terminado"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "CambiosUbicacionBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataCambiosUbicacion 
      Caption         =   "Cambios Ubicacion"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGridCambiosUbicacion 
      Bindings        =   "CambiosUbicacionBatch.frx":08CA
      Height          =   5655
      Left            =   120
      OleObjectBlob   =   "CambiosUbicacionBatch.frx":08ED
      TabIndex        =   4
      ToolTipText     =   "escriba la ubicacion de la tarima"
      Top             =   720
      Width           =   11655
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11040
      Picture         =   "CambiosUbicacionBatch.frx":2877
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "salida"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   495
      Left            =   10080
      Picture         =   "CambiosUbicacionBatch.frx":48E9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "buscar batch"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label LblLinea 
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
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Linea"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Batch"
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
      TabIndex        =   5
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "CambiosUbicacionBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaLinea As Recordset

Private Sub CmdConsultar_Click()
        DataCambiosUbicacion.RecordSource = "Select * From DetalleEntradasProductoTerminado Where Batch = " & TxtTexto.Item(0).Text & " And Linea = '" & TxtTexto.Item(1).Text & "' And Saldo > 0 Order by Tarima"
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

Private Sub Form_Load()
        DataCambiosUbicacion.Connect = GConnect
        DataCambiosUbicacion.DatabaseName = BasedeDatos
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 1 Then
            'BUSCA LINEA
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        TxtTexto.Item(Index).SelStart = 0
        TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
