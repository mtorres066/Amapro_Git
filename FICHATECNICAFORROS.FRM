VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FichaTecnicaForros 
   BackColor       =   &H000000FF&
   Caption         =   "Ficha Tecnica De Forros"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   ControlBox      =   0   'False
   Icon            =   "FichaTecnicaForros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4895
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "FichaTecnicaForros.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameForros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "FichaTecnicaForros.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "FichaTecnicaForros.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -71520
         TabIndex        =   16
         ToolTipText     =   " "
         Top             =   2160
         Width           =   3765
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
         TabIndex        =   15
         Top             =   960
         Width           =   5205
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
      Begin VB.Frame FrameForros 
         Caption         =   "Datos de Forros"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   8115
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoForro"
            DataSource      =   "DataForros"
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataForros"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   840
            Width           =   6915
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FichaTecnicaForros.frx":1488
         Height          =   1905
         Left            =   -74880
         OleObjectBlob   =   "FichaTecnicaForros.frx":14A1
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
         Left            =   -72840
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Data DataForros 
      BackColor       =   &H80000014&
      Caption         =   "Forros"
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
      RecordSource    =   "FichaTecnicaForros"
      Top             =   3960
      Width           =   8175
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6840
      MouseIcon       =   "FichaTecnicaForros.frx":1E94
      Picture         =   "FichaTecnicaForros.frx":22D6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5520
      MouseIcon       =   "FichaTecnicaForros.frx":2718
      Picture         =   "FichaTecnicaForros.frx":2B5A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4200
      MouseIcon       =   "FichaTecnicaForros.frx":308C
      Picture         =   "FichaTecnicaForros.frx":34CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   2880
      MouseIcon       =   "FichaTecnicaForros.frx":3A00
      Picture         =   "FichaTecnicaForros.frx":3E42
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1560
      MouseIcon       =   "FichaTecnicaForros.frx":4374
      Picture         =   "FichaTecnicaForros.frx":47B6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   240
      MouseIcon       =   "FichaTecnicaForros.frx":4CE8
      Picture         =   "FichaTecnicaForros.frx":512A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1200
   End
End
Attribute VB_Name = "FichaTecnicaForros"
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
         FrameForros.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataForros.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameForros.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataForros.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        DataForros.Recordset.AddNew
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
                DataForros.Recordset.Delete
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                DataForros.Recordset.MoveLast
            End If
  
            If DataForros.Recordset.EOF Then
                DataForros.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next
        Bandera = False
        botones
        DataForros.Recordset.CancelUpdate
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
        End If
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        Bandera = True
        botones
        DataForros.Recordset.Edit
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        TxtCod.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   DataForros.Recordset.Update
   
   If Err = 3022 Then
        MsgBox "Codigo de Forro ya Existe", vbOKOnly + vbInformation, "Informacion"
        TxtCod.SetFocus
   ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
   Else
        Bandera = False
        botones
  End If
      
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataForros.RecordSource = ("Select * from FichaTecnicaForros order by " & DBGrid1.Columns(ColIndex).DataField)
    DataForros.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    DataForros.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
    Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
    Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataForros.RecordSource = ("Select * from FichaTecnicaForros where CodigoForro like '" & TxtBuscar.Text & "*'")
            DataForros.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataForros.RecordSource = ("Select * from FichaTecnicaForros where Descripcion like '" & TxtBuscar.Text & "*'")
            DataForros.Refresh
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

