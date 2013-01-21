VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SelloSolve 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento De Sello Solvente"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "SelloSolve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "SelloSolve.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameSelloSolve"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "SelloSolve.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "SelloSolve.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBuscar"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "Lbletiqueta"
      Tab(2).ControlCount=   3
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -72120
         TabIndex        =   16
         ToolTipText     =   " "
         Top             =   2520
         Width           =   2685
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   1200
         Width           =   5205
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   750
            TabIndex        =   15
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   1220
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   2550
            TabIndex        =   14
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameSelloSolve 
         Caption         =   "Datos del Sello Solvente"
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
         Height          =   1335
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   8355
         Begin VB.TextBox TxtCodSelSol 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Codigo"
            DataSource      =   "DataSelloSolve"
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   8
            ToolTipText     =   " "
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtDesSelSol 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataSelloSolve"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   840
            Width           =   7155
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "SelloSolve.frx":1488
         Height          =   2985
         Left            =   -74880
         OleObjectBlob   =   "SelloSolve.frx":14A5
         TabIndex        =   12
         Top             =   720
         Width           =   8385
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
         Left            =   -73440
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Data DataSelloSolve 
      BackColor       =   &H80000014&
      Caption         =   "Sello Solvente"
      Connect         =   "Access"
      DatabaseName    =   "C:\Erick\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SelloSolvente"
      Top             =   4920
      Width           =   8745
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   7320
      MouseIcon       =   "SelloSolve.frx":1E90
      Picture         =   "SelloSolve.frx":22D2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5880
      MouseIcon       =   "SelloSolve.frx":2714
      Picture         =   "SelloSolve.frx":2B56
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4440
      MouseIcon       =   "SelloSolve.frx":3088
      Picture         =   "SelloSolve.frx":34CA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3000
      MouseIcon       =   "SelloSolve.frx":39FC
      Picture         =   "SelloSolve.frx":3E3E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1560
      MouseIcon       =   "SelloSolve.frx":4370
      Picture         =   "SelloSolve.frx":47B2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   120
      MouseIcon       =   "SelloSolve.frx":4CE4
      Picture         =   "SelloSolve.frx":5126
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
End
Attribute VB_Name = "SelloSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim VCodigo As String
Dim VDescripcion As String

Dim RBuscaAmapro As Recordset

Sub botones()
    If Bandera = True Then
         FrameSelloSolve.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodSelSol.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataSelloSolve.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameSelloSolve.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataSelloSolve.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next

        DataSelloSolve.Recordset.AddNew
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        
        Bandera = True
        botones
        TxtCodSelSol.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataSelloSolve.Recordset.Delete
                DataSelloSolve.Recordset.MoveLast
            End If
  
            If DataSelloSolve.Recordset.EOF Then
                DataSelloSolve.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
        
        DataSelloSolve.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        
        Bandera = False
        botones
End Sub

Private Sub CmdEditar_Click()
        
        DataSelloSolve.Recordset.Edit
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        
        Bandera = True
        botones
        
        TxtCodSelSol.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   VCodigo = TxtCodSelSol.Text
   VDescripcion = TxtDesSelSol.Text
   
   DataSelloSolve.Recordset.Update
   
   If Err = 3022 Then
            MsgBox "Codigo de Barniz ya existe", vbOKOnly + vbInformation, "Informacion"
            TxtCodSelSol.SetFocus
   ElseIf Err <> 0 And Err <> 3022 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
   Else
       'BUSCA EL CODIGO EN AMAPRO Y SI LO ENCUENTRA LO MODIFICA SINO LO AGREGA
            Set RBuscaAmapro = Db.OpenRecordset("Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & VCodigo & "'")
                If RBuscaAmapro.RecordCount > 0 Then
                    RBuscaAmapro.Edit
                        RBuscaAmapro!CodigoMateriaPrima = VCodigo
                        RBuscaAmapro!Descripcion = VDescripcion
                    RBuscaAmapro.Update
                Else
                    RBuscaAmapro.AddNew
                        RBuscaAmapro!CodigoMateriaPrima = VCodigo
                        RBuscaAmapro!Descripcion = VDescripcion
                        RBuscaAmapro!Correlativo = 0
                        RBuscaAmapro!Espesor = 0
                        RBuscaAmapro!Minimo = 0
                    RBuscaAmapro.Update
                End If
      Bandera = False
      botones
   End If

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataSelloSolve.RecordSource = ("Select * from SelloSolve order by " & DBGrid1.Columns(ColIndex).DataField)
    DataSelloSolve.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    DataSelloSolve.Connect = GConnect
    DataSelloSolve.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataSelloSolve.RecordSource = ("Select * from SelloSolve where Codigo like '" & TxtBuscar.Text & "*'")
            DataSelloSolve.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataSelloSolve.RecordSource = ("Select * from SelloSolve where Descripcion like '" & TxtBuscar.Text & "*'")
            DataSelloSolve.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub txtcodselsol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtdesselsol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

