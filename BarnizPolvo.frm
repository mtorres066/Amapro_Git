VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form BarnizPolvo 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Tecnica De Barniz Polvo"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ClipControls    =   0   'False
   Icon            =   "BarnizPolvo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "BarnizPolvo.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBarnizPolvo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "BarnizPolvo.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "BarnizPolvo.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).ControlCount=   3
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "BarnizPolvo.frx":1488
         Height          =   3225
         Left            =   -74880
         OleObjectBlob   =   "BarnizPolvo.frx":14A6
         TabIndex        =   12
         Top             =   720
         Width           =   8985
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -73080
         TabIndex        =   16
         ToolTipText     =   " "
         Top             =   2520
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
         Left            =   -74520
         TabIndex        =   13
         Top             =   1320
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
      Begin VB.Frame FrameBarnizPolvo 
         Caption         =   "Datos del Barniz Polvo"
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
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   8835
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataBarnizPolvo"
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Codigo"
            DataSource      =   "DataBarnizPolvo"
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataBarnizPolvo"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   9
            Top             =   840
            Width           =   7635
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   540
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
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   975
         End
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
         Left            =   -74400
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Data DataBarnizPolvo 
      BackColor       =   &H80000014&
      Caption         =   "Barniz Polvo"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BarnizPolvo"
      Top             =   5160
      Width           =   9225
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   7440
      MouseIcon       =   "BarnizPolvo.frx":2035
      Picture         =   "BarnizPolvo.frx":2477
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   6000
      MouseIcon       =   "BarnizPolvo.frx":28B9
      Picture         =   "BarnizPolvo.frx":2CFB
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4560
      MouseIcon       =   "BarnizPolvo.frx":322D
      Picture         =   "BarnizPolvo.frx":366F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3120
      MouseIcon       =   "BarnizPolvo.frx":3BA1
      Picture         =   "BarnizPolvo.frx":3FE3
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1680
      MouseIcon       =   "BarnizPolvo.frx":4515
      Picture         =   "BarnizPolvo.frx":4957
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   240
      MouseIcon       =   "BarnizPolvo.frx":4E89
      Picture         =   "BarnizPolvo.frx":52CB
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " "
      Top             =   4200
      Width           =   1400
   End
End
Attribute VB_Name = "BarnizPolvo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RBuscaAmapro As Recordset

Dim VCodigo As String
Dim VDescripcion As String


Sub botones()
    If Bandera = True Then
         FrameBarnizPolvo.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataBarnizPolvo.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameBarnizPolvo.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataBarnizPolvo.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
  On Error Resume Next
                  
        DataBarnizPolvo.Recordset.AddNew
                
        If Err <> 0 Then
            MsgBox "Error" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
        Else
            Bandera = True
            botones
            TxtCod.SetFocus
            TxtUsuario.Text = GUsuario
        End If
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataBarnizPolvo.Recordset.Delete
                DataBarnizPolvo.Recordset.MoveLast
            End If
  
            If DataBarnizPolvo.Recordset.EOF Then
                DataBarnizPolvo.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next

        DataBarnizPolvo.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
        Else
            Bandera = False
            botones
        End If

        
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next

        DataBarnizPolvo.Recordset.Edit
        
        If Err <> 0 Then
            MsgBox "Error" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
        Else
            Bandera = True
            botones
            TxtCod.SetFocus
            TxtUsuario.Text = GUsuario
        End If
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   VCodigo = TxtCod.Text
   VDescripcion = TxtDes.Text
   
   DataBarnizPolvo.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de el Barniz Polvo ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCod.SetFocus
   ElseIf Err <> 3022 And Err <> 0 Then
      MsgBox "Error" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
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
      
    CmdAgregar.SetFocus
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataBarnizPolvo.RecordSource = ("Select * from BarnizPolvo order by " & DBGrid1.Columns(ColIndex).DataField)
    DataBarnizPolvo.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    DataBarnizPolvo.Connect = GConnect
    DataBarnizPolvo.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataBarnizPolvo.RecordSource = ("Select * from BarnizPolvo where Codigo like '" & TxtBuscar.Text & "*'")
            DataBarnizPolvo.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataBarnizPolvo.RecordSource = ("Select * from BarnizPolvo where Descripcion like '" & TxtBuscar.Text & "*'")
            DataBarnizPolvo.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

