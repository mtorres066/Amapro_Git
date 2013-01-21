VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Bobinas 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bobinas"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "Bobinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabBodegas 
      Height          =   4095
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "Bobinas.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBobinas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Bobinas.frx":0624
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridBobinas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Bobinas.frx":0A76
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar(0)"
      Tab(2).Control(4)=   "CmdBuscar(1)"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "Bobinas.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "Bobinas.frx":11D2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   14
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   1560
         Width           =   1845
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
         TabIndex        =   21
         Top             =   960
         Width           =   2685
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   1220
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   1320
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameBobinas 
         Caption         =   "Datos de Bobina"
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
         Height          =   2775
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   8115
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataBobinas"
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "TipoLamina"
            DataSource      =   "DataBobinas"
            Height          =   285
            Index           =   4
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "Espesor"
            DataSource      =   "DataBobinas"
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   3
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "Diametro"
            DataSource      =   "DataBobinas"
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   2
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoBobina"
            DataSource      =   "DataBobinas"
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataBobinas"
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   6855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Lamina"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Espesor"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Diametro"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGridBobinas 
         Bindings        =   "Bobinas.frx":1614
         Height          =   3105
         Left            =   -74880
         OleObjectBlob   =   "Bobinas.frx":162E
         TabIndex        =   11
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
         Left            =   -70800
         TabIndex        =   22
         Top             =   1560
         Width           =   1935
      End
   End
   Begin VB.Data DataBobinas 
      BackColor       =   &H80000014&
      Caption         =   "Bobinas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Bobinas"
      Top             =   5280
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "Bobinas.frx":26CF
      Picture         =   "Bobinas.frx":2B11
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "Bobinas.frx":2F53
      Picture         =   "Bobinas.frx":3395
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "Bobinas.frx":38C7
      Picture         =   "Bobinas.frx":3D09
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "Bobinas.frx":423B
      Picture         =   "Bobinas.frx":467D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "Bobinas.frx":4BAF
      Picture         =   "Bobinas.frx":4FF1
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "Bobinas.frx":5523
      Picture         =   "Bobinas.frx":5965
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
End
Attribute VB_Name = "Bobinas"
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
         FrameBobinas.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataBobinas.Visible = False
         FrameOpciones.Visible = False
         DBGridBobinas.Visible = False
    Else
         FrameBobinas.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataBobinas.Visible = True
         FrameOpciones.Visible = True
         DBGridBobinas.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataBobinas.Recordset
            If Index = 0 Then
                    'AGREGA UN REGISTRO
                    .AddNew
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(0).SetFocus
                    TxtUsuario.Text = GUsuario
            'EDITAR
            ElseIf Index = 1 Then
                    'EDITA EL REGISTRO
                    .Edit
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(0).SetFocus
                    TxtUsuario.Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "Codigo de Bobina ya existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(0).SetFocus
                      'SI ES CUALQUIER OTRO ERROR
                     ElseIf Err <> 3022 And Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                     End If
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
            'CANCELAR
            ElseIf Index = 3 Then
                    'CANCELA LOS CAMBIOS Y DEJA LOS DATOS COMO ESTABAN
                    .CancelUpdate
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = False
                    botones
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        DataBobinas.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataBobinas.Recordset.MoveLast
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataBobinas.Recordset.EOF Then
                        DataBobinas.Recordset.MoveLast
                        If Err = 3021 Then
                            mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                        End If
                    End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        End With
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    With DataBobinas
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                .RecordSource = ("Select * from Bobinas where CodigoBobina like '" & TxtBuscar.Text & "*'")
                .Refresh
                DBGridBobinas.Refresh
            ElseIf OptNombre.Value = True Then
                .RecordSource = ("Select * from Bobinas where Descripcion like '" & TxtBuscar.Text & "*'")
                .Refresh
                DBGridBobinas.Refresh
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From Bobinas"
                .Refresh
                DBGridBobinas.Refresh
        End If
    End With
        TabBodegas.Tab = 1
End Sub

Private Sub dbgridbobinas_HeadClick(ByVal ColIndex As Integer)
    DataBobinas.RecordSource = ("Select * from Bobinas order by " & DBGridBobinas.Columns(ColIndex).DataField)
    DataBobinas.Refresh
    DBGridBobinas.Refresh
End Sub

Private Sub Form_Load()
    DataBobinas.Connect = GConnect
    DataBobinas.DatabaseName = BasedeDatos
End Sub

Private Sub OptCodigo_Click()
Lbletiqueta.Caption = "Codigo"
End Sub

Private Sub OptNombre_Click()
Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
