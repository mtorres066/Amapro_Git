VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EmpleadosAumentos 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aumentos De Salario a Empleados"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "EmpleadosAumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabPuestos 
      Height          =   4215
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "EmpleadosAumentos.frx":2E7A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "EmpleadosAumentos.frx":3194
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "EmpleadosAumentos.frx":35E6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBuscar"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdBuscar(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdBuscar(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Lbletiqueta"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   12
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "EmpleadosAumentos.frx":3A38
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "EmpleadosAumentos.frx":3D42
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   2055
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
         TabIndex        =   19
         Top             =   960
         Width           =   2805
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   1080
            TabIndex        =   11
            ToolTipText     =   " "
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos De La Falta"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Usuario"
            DataSource      =   "Data"
            Height          =   285
            Index           =   2
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Descripcion"
            DataSource      =   "Data"
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   6855
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Codigo"
            DataSource      =   "Data"
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   840
         End
      End
      Begin MSDBGrid.DBGrid DBGrid 
         Bindings        =   "EmpleadosAumentos.frx":4184
         Height          =   3345
         Left            =   -74880
         OleObjectBlob   =   "EmpleadosAumentos.frx":4197
         TabIndex        =   9
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
         TabIndex        =   20
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Data Data 
      BackColor       =   &H80000014&
      Caption         =   "Faltas"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EmpleadosFaltas"
      Top             =   5160
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "EmpleadosAumentos.frx":4D25
      Picture         =   "EmpleadosAumentos.frx":5167
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1320
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "EmpleadosAumentos.frx":71D9
      Picture         =   "EmpleadosAumentos.frx":761B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "EmpleadosAumentos.frx":7B4D
      Picture         =   "EmpleadosAumentos.frx":7F8F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "EmpleadosAumentos.frx":84C1
      Picture         =   "EmpleadosAumentos.frx":8903
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "EmpleadosAumentos.frx":8E35
      Picture         =   "EmpleadosAumentos.frx":9277
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "EmpleadosAumentos.frx":97A9
      Picture         =   "EmpleadosAumentos.frx":9BEB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1200
   End
End
Attribute VB_Name = "EmpleadosAumentos"
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
         FramePuestos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         Data.Visible = False
         FrameOpciones.Visible = False
         DBGrid.Visible = False
    Else
         FramePuestos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         Data.Visible = True
         FrameOpciones.Visible = True
         DBGrid.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With Data.Recordset
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
                    Txttexto.Item(0).SetFocus
                    Txttexto.Item(2).Text = GUsuario
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
                    Txttexto.Item(0).SetFocus
                    Txttexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                   
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "Codigo De La Falta Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(0).SetFocus
                        Exit Sub
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
                        Data.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        Data.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If Data.Recordset.EOF Then
                        Data.Recordset.MoveLast
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
    With Data
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                .RecordSource = ("Select * from EmpleadosFaltas where Codigo Like '*" & TxtBuscar.Text & "*'")
                .Refresh
                DBGrid.Refresh
            ElseIf OptDescripcion.Value = True Then
                .RecordSource = ("Select * from EmpleadosFaltas where Descripcion Like '*" & TxtBuscar.Text & "*'")
                .Refresh
                DBGrid.Refresh
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From EmpleadosFaltas"
                .Refresh
                DBGrid.Refresh
        End If
    End With
        TabPuestos.Tab = 1
End Sub


Private Sub DBGrid_HeadClick(ByVal ColIndex As Integer)
    Data.RecordSource = ("Select * from EmpleadosFaltas order by " & DBGrid.Columns(ColIndex).DataField)
    Data.Refresh
    DBGrid.Refresh
End Sub


Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        Data.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        Data.DatabaseName = BasedeDatos

End Sub

Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = "Codigo"
End Sub

Private Sub OptDescripcion_Click()
        Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_GotFocus()
        TxtBuscar.SelStart = 0
        TxtBuscar.SelLength = Len(TxtBuscar.Text)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        Txttexto.Item(Index).SelStart = 0
        Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
