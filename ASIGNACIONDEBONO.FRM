VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AsignaciondeBono 
   Caption         =   "Asignacion de BONO de Grupos"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "AsignaciondeBono.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameConsultas 
      Height          =   7335
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Data DataConsultas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10800
         Picture         =   "AsignaciondeBono.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "AsignaciondeBono.frx":0614
         Height          =   6975
         Left            =   120
         OleObjectBlob   =   "AsignaciondeBono.frx":0630
         TabIndex        =   20
         Top             =   240
         Width           =   10575
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "AsignaciondeBono.frx":100B
      Height          =   4905
      Left            =   120
      OleObjectBlob   =   "AsignaciondeBono.frx":1028
      TabIndex        =   17
      Top             =   2400
      Width           =   11505
   End
   Begin VB.Data DataAsignacion 
      BackColor       =   &H80000014&
      Caption         =   "Asignacion de Variables a Grupos"
      Connect         =   "Access"
      DatabaseName    =   "C:\DPI\DPI.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AsignaciondeBono"
      Top             =   7440
      Width           =   11865
   End
   Begin VB.Frame FrameOpciones 
      Caption         =   "Opciones de Busqueda"
      Height          =   855
      Left            =   6300
      TabIndex        =   16
      Top             =   40
      Width           =   5205
      Begin VB.OptionButton OptNombre 
         Caption         =   "Mes Y Año"
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         ToolTipText     =   " "
         Top             =   360
         Width           =   1340
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Grupo"
         Height          =   225
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   " "
         Top             =   360
         Value           =   -1  'True
         Width           =   1220
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "AsignaciondeBono.frx":19FB
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   7800
      TabIndex        =   11
      ToolTipText     =   " "
      Top             =   1080
      Width           =   3765
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   705
      Left            =   9720
      MouseIcon       =   "AsignaciondeBono.frx":1F7D
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":23BF
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1785
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   705
      Left            =   7800
      MouseIcon       =   "AsignaciondeBono.frx":26C9
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":2B0B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1900
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   705
      Left            =   5880
      MouseIcon       =   "AsignaciondeBono.frx":2F4D
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":338F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   705
      Left            =   3960
      MouseIcon       =   "AsignaciondeBono.frx":3919
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":3D5B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1900
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   705
      Left            =   2040
      MouseIcon       =   "AsignaciondeBono.frx":419D
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":45DF
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1900
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   705
      Left            =   120
      MouseIcon       =   "AsignaciondeBono.frx":4A21
      MousePointer    =   99  'Custom
      Picture         =   "AsignaciondeBono.frx":4E63
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   1560
      Width           =   1900
   End
   Begin VB.Frame FrameAsignacion 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   150
      TabIndex        =   12
      Top             =   0
      Width           =   6075
      Begin MSMask.MaskEdBox MskBon 
         DataField       =   "Bono"
         DataSource      =   "DataAsignacion"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtMesAño 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "Mes"
         DataSource      =   "DataAsignacion"
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   1
         ToolTipText     =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtCodGru 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         DataField       =   "GrupoAsignacionBono"
         DataSource      =   "DataAsignacion"
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblPeso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label LblVariable 
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label LblGrupo 
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Bono"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Mes y Año"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Lbletiqueta 
      Caption         =   "Grupo"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "AsignaciondeBono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BGrupos As Boolean
Dim BVariables As Boolean

Dim RGrupos As Recordset
Dim RVariables As Recordset

Dim RTotalVariables As Recordset
Dim RSumavariables As Recordset

Dim VTotalVariables As Currency
Dim VSumaVariables As Currency

Dim VPuntos As Currency


Sub botones()
    If Bandera = True Then
         FrameAsignacion.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodGru.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataAsignacion.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameAsignacion.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataAsignacion.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
        Bandera = True
        botones
        DataAsignacion.Recordset.AddNew
        TxtCodGru.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataAsignacion.Recordset.Delete
                DataAsignacion.Recordset.MoveLast
            End If
  
            If DataAsignacion.Recordset.EOF Then
                DataAsignacion.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        DataAsignacion.Recordset.CancelUpdate
End Sub

Private Sub CmdEditar_Click()
        Bandera = True
        botones
        DataAsignacion.Recordset.Edit
        TxtCodGru.SetFocus
        VPuntos = MskBon.Text
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
          
   'GRABA DATOS
   DataAsignacion.Recordset.Update
   
   If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        TxtCodGru.SetFocus
   Else
      Bandera = False
      botones
  End If
      
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataAsignacion.RecordSource = ("Select * from AsignacionDeBono order by " & DBGrid1.Columns(ColIndex).DataField)
    DataAsignacion.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub DBGridConsultas_DblClick()
            If BGrupos = True Then
                FrameConsultas.Visible = False
                TxtCodGru.Text = DBGridConsultas.Columns(0)
                TxtCodGru.SetFocus
            End If
            If BVariables = True Then
                FrameConsultas.Visible = False
                TxtMesAño.Text = DBGridConsultas.Columns(0)
                TxtMesAño.SetFocus
            End If
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then
            If BGrupos = True Then
                FrameConsultas.Visible = False
                TxtCodGru.Text = DBGridConsultas.Columns(0)
                TxtCodGru.SetFocus
            End If
            If BVariables = True Then
                FrameConsultas.Visible = False
                TxtMesAño.Text = DBGridConsultas.Columns(0)
                TxtMesAño.SetFocus
            End If
        End If

End Sub

Private Sub Form_Load()
    DataAsignacion.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub


Private Sub MskBon_GotFocus()
    MskBon.SelStart = 0
    MskBon.SelLength = Len(MskBon.Text)
End Sub

Private Sub MskBon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Grupo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Mes y Año"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataAsignacion.RecordSource = ("Select * from AsignacionDeBono where GrupoAsignacionBono like '" & TxtBuscar.Text & "*'")
            DataAsignacion.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataAsignacion.RecordSource = ("Select * from AsignacionDeBono where Mes like '" & TxtBuscar.Text & "*'")
            DataAsignacion.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub TxtCodGru_Change()
    Set RGrupos = Db.OpenRecordset("Select DescripcionGrupo From Grupos Where CodigoGrupo = '" & TxtCodGru.Text & "'")
    If RGrupos.RecordCount > 0 Then
        LblGrupo.Caption = RGrupos(0)
    Else
        LblGrupo.Caption = ""
    End If

End Sub

Private Sub TxtCodGru_DblClick()
    BGrupos = True
    BVariables = False
    DataConsultas.RecordSource = ("Select * from Grupos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus

End Sub

Private Sub TxtCodGru_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BGrupos = True
    BVariables = False
    DataConsultas.RecordSource = ("Select * from Grupos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub

Private Sub TxtDesAre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub



Private Sub TxtMesAño_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub
