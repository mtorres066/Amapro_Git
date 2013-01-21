VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Tolerancia 
   Caption         =   "Tolerancia"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Tolerancia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Tolerancia.frx":08CA
      Height          =   5385
      Left            =   120
      OleObjectBlob   =   "Tolerancia.frx":08E7
      TabIndex        =   16
      Top             =   1920
      Width           =   11505
   End
   Begin VB.Data DataTolerancia 
      BackColor       =   &H80000014&
      Caption         =   "Tolerancia"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tolerancia"
      Top             =   7440
      Width           =   11865
   End
   Begin VB.Frame FrameOpciones 
      Caption         =   "Opciones de Busqueda"
      Height          =   740
      Left            =   6300
      TabIndex        =   15
      Top             =   40
      Width           =   5205
      Begin VB.OptionButton OptNombre 
         Caption         =   "&Margen"
         Height          =   195
         Left            =   2550
         TabIndex        =   9
         ToolTipText     =   " "
         Top             =   300
         Width           =   1340
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "&Tolerancia"
         Height          =   225
         Left            =   750
         TabIndex        =   8
         ToolTipText     =   " "
         Top             =   300
         Value           =   -1  'True
         Width           =   1220
      End
   End
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   7650
      TabIndex        =   10
      ToolTipText     =   " "
      Top             =   900
      Width           =   3765
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   465
      Left            =   9900
      MouseIcon       =   "Tolerancia.frx":12BA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1785
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   465
      Left            =   7950
      MouseIcon       =   "Tolerancia.frx":16FC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1900
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   6000
      MouseIcon       =   "Tolerancia.frx":1B3E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   4050
      MouseIcon       =   "Tolerancia.frx":1F80
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1900
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   465
      Left            =   2100
      MouseIcon       =   "Tolerancia.frx":23C2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1900
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   465
      Left            =   150
      MouseIcon       =   "Tolerancia.frx":2804
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   1350
      Width           =   1900
   End
   Begin VB.Frame FrameTolerancia 
      Caption         =   "Datos de la Tolerancia"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   150
      TabIndex        =   11
      Top             =   0
      Width           =   6075
      Begin VB.TextBox TxtDesTol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "MARGEN"
         DataSource      =   "DataTolerancia"
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   " "
         Top             =   840
         Width           =   4875
      End
      Begin VB.TextBox TxtCodTol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "TOLERAN"
         DataSource      =   "DataTolerancia"
         Height          =   285
         Left            =   1050
         TabIndex        =   0
         ToolTipText     =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Margen"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tolerancia"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Lbletiqueta 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   6300
      TabIndex        =   14
      Top             =   900
      Width           =   1215
   End
End
Attribute VB_Name = "Tolerancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Sub botones()
    If Bandera = True Then
         FrameTolerancia.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodTol.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataTolerancia.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameTolerancia.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataTolerancia.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
        Bandera = True
        botones
        DataTolerancia.Recordset.AddNew
        TxtCodTol.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataTolerancia.Recordset.Delete
                DataTolerancia.Recordset.MoveLast
            End If
  
            If DataTolerancia.Recordset.EOF Then
                DataTolerancia.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        DataTolerancia.Recordset.CancelUpdate
End Sub

Private Sub CmdEditar_Click()
        Bandera = True
        botones
        DataTolerancia.Recordset.Edit
        TxtCodTol.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   DataTolerancia.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Tolerancia ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCodTol.SetFocus
   Else
      Bandera = False
      botones
  End If
      
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataTolerancia.RecordSource = ("Select * from Tolerancia order by " & DBGrid1.Columns(ColIndex).DataField)
    DataTolerancia.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    DataTolerancia.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataTolerancia.RecordSource = ("Select * from Tolerancia where Toleran like '" & TxtBuscar.Text & "*'")
            DataTolerancia.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataTolerancia.RecordSource = ("Select * from Tolerancia where Margen like '" & TxtBuscar.Text & "*'")
            DataTolerancia.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub txtcodtol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtDestol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

