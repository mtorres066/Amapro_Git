VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Advertencias 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo De Advertencias Supervision De Materias Primas"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "Advertencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "Advertencias.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAdvertencias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Advertencias.frx":075C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Advertencias.frx":0BAE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBuscar"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "Lbletiqueta"
      Tab(2).ControlCount=   3
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -69840
         TabIndex        =   11
         ToolTipText     =   " "
         Top             =   2640
         Width           =   2805
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones de Busqueda"
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
         Height          =   740
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   4245
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   750
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   1220
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   2550
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameAdvertencias 
         Caption         =   "Datos de Advertencias"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   8115
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoAdvertencia"
            DataSource      =   "DataAdvertencias"
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataAdvertencias"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   1200
            Width           =   6855
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Advertencias.frx":1000
         Height          =   2625
         Left            =   -74880
         OleObjectBlob   =   "Advertencias.frx":101F
         TabIndex        =   8
         Top             =   720
         Width           =   8145
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -71520
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.Data DataAdvertencias 
      BackColor       =   &H80000014&
      Caption         =   "Advertencias De Materias Primas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Advertencias"
      Top             =   4560
      Width           =   8115
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6840
      MouseIcon       =   "Advertencias.frx":1A22
      Picture         =   "Advertencias.frx":1E64
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5520
      MouseIcon       =   "Advertencias.frx":22A6
      Picture         =   "Advertencias.frx":26E8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4200
      MouseIcon       =   "Advertencias.frx":2C1A
      Picture         =   "Advertencias.frx":305C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   2880
      MouseIcon       =   "Advertencias.frx":358E
      Picture         =   "Advertencias.frx":39D0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1560
      MouseIcon       =   "Advertencias.frx":3F02
      Picture         =   "Advertencias.frx":4344
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   240
      MouseIcon       =   "Advertencias.frx":4876
      Picture         =   "Advertencias.frx":4CB8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   3600
      Width           =   1200
   End
End
Attribute VB_Name = "Advertencias"
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
         FrameAdvertencias.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         dataadvertencias.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameAdvertencias.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         dataadvertencias.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        dataadvertencias.Recordset.AddNew
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        TxtCod.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                dataadvertencias.Recordset.Delete
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                    Exit Sub
                End If
                dataadvertencias.Recordset.MoveLast
            End If
  
            If dataadvertencias.Recordset.EOF Then
                dataadvertencias.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next
        Bandera = False
        botones
        dataadvertencias.Recordset.CancelUpdate
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        Bandera = True
        botones
        dataadvertencias.Recordset.Edit
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        TxtCod.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   dataadvertencias.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Bodega ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCod.SetFocus
   ElseIf Err <> 3022 And Err <> 0 Then
      MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
      Exit Sub
   Else
      Bandera = False
      botones
      CmdAgregar.SetFocus
  End If
      
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    dataadvertencias.RecordSource = ("Select * from BodegasMateriaPrima order by " & DBGrid1.Columns(ColIndex).DataField)
    dataadvertencias.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    dataadvertencias.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub



Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            dataadvertencias.RecordSource = ("Select * from BodegasMateriaPrima where CodigoBodega like '" & TxtBuscar.Text & "*'")
            dataadvertencias.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            dataadvertencias.RecordSource = ("Select * from BodegasMateriaPrima where Descripcion like '" & TxtBuscar.Text & "*'")
            dataadvertencias.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub TxtCod_GotFocus()
    TxtCod.SelStart = 0
    TxtCod.SelLength = Len(TxtCod.Text)
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtDes_GotFocus()
    TxtDes.SelStart = 0
    TxtDes.SelLength = Len(TxtDes.Text)
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

