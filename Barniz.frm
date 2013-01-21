VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Barniz 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento De Barniz"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "Barniz.frx":0000
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
      TabPicture(0)   =   "Barniz.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBarniz"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Barniz.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Barniz.frx":1036
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
      Begin VB.Frame FrameBarniz 
         Caption         =   "Datos del Barniz"
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
         Height          =   1695
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   8355
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataBarniz"
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox TxtCodBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "BARNIZ"
            DataSource      =   "DataBarniz"
            Height          =   285
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtDesBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "DESCRIP"
            DataSource      =   "DataBarniz"
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   9
            Top             =   840
            Width           =   7155
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
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Barniz.frx":1488
         Height          =   2985
         Left            =   -74880
         OleObjectBlob   =   "Barniz.frx":14A1
         TabIndex        =   12
         Top             =   720
         Width           =   8505
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
   Begin VB.Data DataBarniz 
      BackColor       =   &H80000014&
      Caption         =   "Barniz"
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
      RecordSource    =   "Barniz"
      Top             =   4920
      Width           =   8745
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   7320
      MouseIcon       =   "Barniz.frx":202C
      Picture         =   "Barniz.frx":246E
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
      MouseIcon       =   "Barniz.frx":28B0
      Picture         =   "Barniz.frx":2CF2
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
      MouseIcon       =   "Barniz.frx":3224
      Picture         =   "Barniz.frx":3666
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
      MouseIcon       =   "Barniz.frx":3B98
      Picture         =   "Barniz.frx":3FDA
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
      MouseIcon       =   "Barniz.frx":450C
      Picture         =   "Barniz.frx":494E
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
      MouseIcon       =   "Barniz.frx":4E80
      Picture         =   "Barniz.frx":52C2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " "
      Top             =   3960
      Width           =   1400
   End
End
Attribute VB_Name = "Barniz"
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
         FrameBarniz.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodBar.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataBarniz.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameBarniz.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataBarniz.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next

        
        DataBarniz.Recordset.AddNew
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
                
        Bandera = True
        botones
        TxtCodBar.SetFocus
        TxtUsuario.Text = GUsuario
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataBarniz.Recordset.Delete
                DataBarniz.Recordset.MoveLast
            End If
  
            If DataBarniz.Recordset.EOF Then
                DataBarniz.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next

        DataBarniz.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        
        Bandera = False
        botones
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next

        DataBarniz.Recordset.Edit
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
            
        End If
        
        Bandera = True
        botones
        TxtCodBar.SetFocus
        TxtUsuario.Text = GUsuario
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   DataBarniz.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Barniz ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCodBar.SetFocus
   ElseIf Err <> 0 And Err <> 3022 Then
      MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
      Exit Sub
   Else
      Bandera = False
      botones
  End If
      CmdAgregar.SetFocus
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataBarniz.RecordSource = ("Select * from Barniz order by " & DBGrid1.Columns(ColIndex).DataField)
    DataBarniz.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub Form_Load()
    DataBarniz.Connect = GConnect
    DataBarniz.DatabaseName = BasedeDatos
End Sub


Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_Change()
        
        If OptCodigo.Value = True Then
            DataBarniz.RecordSource = ("Select * from Barniz where Barniz like '" & TxtBuscar.Text & "*'")
            DataBarniz.Refresh
            DBGrid1.Refresh
        ElseIf OptNombre.Value = True Then
            DataBarniz.RecordSource = ("Select * from Barniz where Descrip like '" & TxtBuscar.Text & "*'")
            DataBarniz.Refresh
            DBGrid1.Refresh
        End If
        
End Sub

Private Sub TxtCodBar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtDesbar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

