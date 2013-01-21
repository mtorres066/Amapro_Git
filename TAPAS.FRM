VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Tapas 
   Caption         =   "Ficha Tecnica De Tapas"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Tapas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   3855
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10680
         Picture         =   "Tapas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataConsultas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "Tapas.frx":0BD4
         Height          =   3495
         Left            =   120
         OleObjectBlob   =   "Tapas.frx":0BF0
         TabIndex        =   32
         Top             =   240
         Width           =   10455
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Tapas.frx":15CB
      Height          =   3345
      Left            =   120
      OleObjectBlob   =   "Tapas.frx":15E3
      TabIndex        =   22
      Top             =   3960
      Width           =   11505
   End
   Begin VB.Data DataTapas 
      BackColor       =   &H80000014&
      Caption         =   "Tapas"
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
      RecordSource    =   "Tapas"
      Top             =   7440
      Width           =   11865
   End
   Begin VB.Frame FrameOpciones 
      Caption         =   "Opciones de Busqueda"
      Height          =   740
      Left            =   6300
      TabIndex        =   21
      Top             =   40
      Width           =   5205
      Begin VB.OptionButton OptNombre 
         Caption         =   "&Descripcion"
         Height          =   195
         Left            =   2550
         TabIndex        =   15
         ToolTipText     =   " "
         Top             =   300
         Width           =   1340
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "&Codigo"
         Height          =   225
         Left            =   750
         TabIndex        =   14
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
      TabIndex        =   16
      ToolTipText     =   " "
      Top             =   900
      Width           =   3765
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   465
      Left            =   9720
      MouseIcon       =   "Tapas.frx":1FB6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1785
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   465
      Left            =   7800
      MouseIcon       =   "Tapas.frx":23F8
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1900
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   5880
      MouseIcon       =   "Tapas.frx":283A
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   3960
      MouseIcon       =   "Tapas.frx":2C7C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1900
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   465
      Left            =   2040
      MouseIcon       =   "Tapas.frx":30BE
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1900
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   465
      Left            =   120
      MouseIcon       =   "Tapas.frx":3500
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   " "
      Top             =   3360
      Width           =   1900
   End
   Begin VB.Frame FrameTapas 
      Caption         =   "Datos de la Tapa"
      Enabled         =   0   'False
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6075
      Begin VB.TextBox TxtBar 
         Appearance      =   0  'Flat
         DataField       =   "BARNIZ"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox TxtDur 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "DUREZA"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtDia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "DIAMETRO"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtEsp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ESPESOR"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TxtAlt 
         Appearance      =   0  'Flat
         DataField       =   "ALTERNO"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtFor 
         Appearance      =   0  'Flat
         DataField       =   "FORMA"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtDesTap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "DESCRIP"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   " "
         Top             =   720
         Width           =   4875
      End
      Begin VB.TextBox TxtCodTap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "TAPA"
         DataSource      =   "DataTapas"
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   " "
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LblBarniz 
         Height          =   255
         Left            =   2400
         TabIndex        =   29
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "Barniz"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Dureza"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Diametro"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Espesor"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Alterno"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Forma"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Descipcion"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Lbletiqueta 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   6300
      TabIndex        =   20
      Top             =   900
      Width           =   1215
   End
End
Attribute VB_Name = "Tapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim Mensaje As String
Dim Buscar As String

Dim RBarniz As Recordset

Sub botones()
    If Bandera = True Then
         FrameTapas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodTap.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataTapas.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameTapas.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataTapas.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
        Bandera = True
        botones
        DataTapas.Recordset.AddNew
        TxtCodTap.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If Mensaje = vbOK Then
                DataTapas.Recordset.Delete
                DataTapas.Recordset.MoveLast
            End If
  
            If DataTapas.Recordset.EOF Then
                DataTapas.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        DataTapas.Recordset.CancelUpdate
End Sub

Private Sub CmdEditar_Click()
        Bandera = True
        botones
        DataTapas.Recordset.Edit
        TxtCodTap.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   DataTapas.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Tapa ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCodTap.SetFocus
   Else
      Bandera = False
      botones
  End If
      
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataTapas.RecordSource = ("Select * from Tapas order by " & DBGrid1.Columns(ColIndex).DataField)
    DataTapas.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub DBGridConsultas_DblClick()
    TxtBar.Text = DBGridConsultas.Columns(0)
    TxtBar.SetFocus
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    
If KeyAscii = 27 Then
    TxtBar.Text = DBGridConsultas.Columns(0)
    TxtBar.SetFocus
    FrameConsultas.Visible = False
End If
End Sub


Private Sub Form_Load()
    DataTapas.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtAlt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub


Private Sub TxtBar_Change()
    Set RBarniz = Db.OpenRecordset("Select Descrip From Barniz Where Barniz = '" & TxtBar.Text & "'")
    If RBarniz.RecordCount > 0 Then
        LblBarniz.Caption = RBarniz(0)
    Else
        LblBarniz.Caption = ""
    End If
End Sub

Private Sub TxtBar_DblClick()
    DataConsultas.RecordSource = ("Select * from Barniz")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End Sub

Private Sub TxtBar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    DataConsultas.RecordSource = ("Select * from Barniz")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub


Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataTapas.RecordSource = ("Select * from Tapas where Tapa like '" & TxtBuscar.Text & "*'")
            DataTapas.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataTapas.RecordSource = ("Select * from Tapas where Descrip like '" & TxtBuscar.Text & "*'")
            DataTapas.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub TxtCodTap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtDestap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDur_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtEsp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtFor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub
