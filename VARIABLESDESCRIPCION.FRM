VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form VariablesDescripcion 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encabezado De Catalogos"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8025
   Icon            =   "VariablesDescripcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "VariablesDescripcion.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameVariablesDescripcion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "VariablesDescripcion.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameVariablesDescripcion 
         Caption         =   "Datos de Catalogo"
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
         Left            =   840
         TabIndex        =   9
         Top             =   1680
         Width           =   6075
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoVariable"
            DataSource      =   "DataVariablesDescripcion"
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "DescripcionVariable"
            DataSource      =   "DataVariablesDescripcion"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   840
            Width           =   4875
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "VariablesDescripcion.frx":1036
         Height          =   3465
         Left            =   -74880
         OleObjectBlob   =   "VariablesDescripcion.frx":105D
         TabIndex        =   8
         Top             =   720
         Width           =   7545
      End
   End
   Begin VB.Data DataVariablesDescripcion 
      BackColor       =   &H80000014&
      Caption         =   "Encabezado De Catalogos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VariablesDescripcion"
      Top             =   5280
      Width           =   7785
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   600
      Left            =   6720
      MouseIcon       =   "VariablesDescripcion.frx":1A68
      Picture         =   "VariablesDescripcion.frx":1EAA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   600
      Left            =   5400
      MouseIcon       =   "VariablesDescripcion.frx":22EC
      Picture         =   "VariablesDescripcion.frx":272E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   600
      Left            =   4080
      MouseIcon       =   "VariablesDescripcion.frx":2C60
      Picture         =   "VariablesDescripcion.frx":30A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   600
      Left            =   2760
      MouseIcon       =   "VariablesDescripcion.frx":35D4
      Picture         =   "VariablesDescripcion.frx":3A16
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   600
      Left            =   1440
      MouseIcon       =   "VariablesDescripcion.frx":3F48
      Picture         =   "VariablesDescripcion.frx":438A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   600
      Left            =   120
      MouseIcon       =   "VariablesDescripcion.frx":48BC
      Picture         =   "VariablesDescripcion.frx":4CFE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   4560
      Width           =   1200
   End
End
Attribute VB_Name = "VariablesDescripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim Mensaje As String
Dim Buscar As String

Sub botones()
    If Bandera = True Then
         FrameVariablesDescripcion.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         DataVariablesDescripcion.Visible = False
         DBGrid1.Visible = False
    Else
         FrameVariablesDescripcion.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         DataVariablesDescripcion.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        'AGREGA UN REGISTRO
        DataVariablesDescripcion.Recordset.AddNew
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
            Bandera = True
            botones
            TxtCod.SetFocus
        End If
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If Mensaje = vbOK Then
                DataVariablesDescripcion.Recordset.Delete
                DataVariablesDescripcion.Recordset.MoveLast
            End If
  
            If DataVariablesDescripcion.Recordset.EOF Then
                DataVariablesDescripcion.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
    On Error Resume Next
        'CANCELA Y NO GRABA LOS CAMBIOS
        DataVariablesDescripcion.Recordset.CancelUpdate
        If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
                Bandera = False
                botones
        End If
End Sub

Private Sub CmdEditar_Click()
    On Error Resume Next
        'EDITA EL REGISTRO
        DataVariablesDescripcion.Recordset.Edit
        If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
                Bandera = True
                botones
                TxtCod.SetFocus
        End If
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   'GRABA DATOS
   DataVariablesDescripcion.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Variable ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCod.SetFocus
   ElseIf Err <> 3022 And Err <> 0 Then
      MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
   Else
      Bandera = False
      botones
  End If

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataVariablesDescripcion.RecordSource = ("Select * from VariablesDescripcion order by " & DBGrid1.Columns(ColIndex).DataField)
    DataVariablesDescripcion.Refresh
    DBGrid1.Refresh
End Sub

Private Sub Form_Load()
    DataVariablesDescripcion.DatabaseName = BasedeDatos
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

