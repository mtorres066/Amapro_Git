VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Defectos 
   BackColor       =   &H000000FF&
   Caption         =   "Defectos"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "Defectos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "Defectos.frx":08CA
      Picture         =   "Defectos.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Primer Registro"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "Defectos.frx":123E
      Picture         =   "Defectos.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Registro Anterior"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "Defectos.frx":1BB2
      Picture         =   "Defectos.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Siguiente Registro"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7800
      MouseIcon       =   "Defectos.frx":2526
      Picture         =   "Defectos.frx":2968
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ultimo Registro"
      Top             =   4440
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   4215
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "Defectos.frx":2E9A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Defectos.frx":31B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Defectos.frx":3606
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBuscar"
      Tab(2).Control(1)=   "CmdBuscar(1)"
      Tab(2).Control(2)=   "CmdBuscar(0)"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "Lbletiqueta"
      Tab(2).ControlCount=   5
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Defecto"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descrip"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4424.882
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   17
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "Defectos.frx":3A58
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "Defectos.frx":3D62
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   15
         Top             =   960
         Width           =   2805
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   1080
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos Del Defecto"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Index           =   3
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   27
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   2
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label3 
            Caption         =   "Tipos De Defecto ( 0 = menor    1 = mayor    2 = critico)"
            Height          =   255
            Left            =   3360
            TabIndex        =   26
            Top             =   1080
            Width           =   4575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   840
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
         Left            =   -70800
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "Defectos.frx":41A4
      Picture         =   "Defectos.frx":45E6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "Defectos.frx":6658
      Picture         =   "Defectos.frx":6A9A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "Defectos.frx":6FCC
      Picture         =   "Defectos.frx":740E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3120
      MouseIcon       =   "Defectos.frx":7940
      Picture         =   "Defectos.frx":7D82
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2040
      MouseIcon       =   "Defectos.frx":82B4
      Picture         =   "Defectos.frx":86F6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "Defectos.frx":8C28
      Picture         =   "Defectos.frx":906A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1000
   End
End
Attribute VB_Name = "Defectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RDefectos As New ADODB.Recordset
Dim BEditar As Boolean


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
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
         CmdBuscar.Item(0).Visible = False
         CmdBuscar.Item(1).Visible = False
        
         FrameOpciones.Visible = False
         DataGrid1.Visible = False
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
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         
         CmdBuscar.Item(0).Visible = True
         CmdBuscar.Item(1).Visible = True

         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
            'AGREGAR
            If Index = 0 Then
                    Bandera = True
                    botones
                    Limpia_Campos
                    Txttexto.Item(0).Enabled = True
                    Txttexto.Item(0).SetFocus
                    Txttexto.Item(3).Text = GUsuario
                    BEditar = False
            'EDITAR
            ElseIf Index = 1 Then
                    Bandera = True
                    botones
                    Txttexto.Item(0).Enabled = False
                    Txttexto.Item(1).SetFocus
                    Txttexto.Item(3).Text = GUsuario
                    BEditar = True
            'GRABAR
            ElseIf Index = 2 Then
                    If Txttexto.Item(2).Text <> "0" And Txttexto.Item(2).Text <> "1" And Txttexto.Item(2).Text <> "2" Then
                         MsgBox "Tipo De Defecto Incorrecto", vbOKOnly + vbInformation, "Informacion"
                         Txttexto.Item(2).SetFocus
                         Exit Sub
                    End If

            
            
                    If BEditar = False Then 'ESTA AGREGANDO UN REGISTRO
                         Conexion.Execute "Insert Into Defectos Values('" & Txttexto.Item(0).Text & "', '" & Txttexto.Item(1).Text & "', '" & Txttexto.Item(2).Text & "', '" & Txttexto.Item(3).Text & "')"
                    Else 'ESTA EDITANDO UN REGISTRO
                         Conexion.Execute "UPDATE Defectos SET Descrip = '" & Txttexto.Item(1).Text & "', Tipo = '" & Txttexto.Item(2).Text & "', Usuario = '" & Txttexto.Item(3).Text & "' Where Defecto = '" & Txttexto.Item(0).Text & "'"
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Defecto Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            Txttexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Defecto Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            Txttexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                        
                        
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                        Txttexto.Item(0).Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RDefectos.Requery
                        RDefectos.MoveLast
                        Llena_Campos
            'CANCELAR
            ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    Txttexto.Item(0).Enabled = True
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RDefectos.Delete
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RDefectos.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDefectos.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                        
                    End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RDefectos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RDefectos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RDefectos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RDefectos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RDefectos.BOF Then
        RDefectos.MoveFirst
    ElseIf RDefectos.EOF Then
        RDefectos.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RDefectos = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from Defectos where Defecto Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from Defectos where UPPER(Defecto) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from Defectos where Descrip Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from Defectos where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RDefectos, "Select * From Defectos")
        End If
        
        Set DataGrid1.DataSource = RDefectos
    
        TabPuestos.Tab = 1
End Sub


Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
                RDefectos.Sort = RDefectos.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If
    
End Sub


Private Sub Form_Load()
        Set RDefectos = New ADODB.Recordset
        Call Abrir_Recordset(RDefectos, "Select * From Defectos")
        Set DataGrid1.DataSource = RDefectos
        Llena_Campos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        RDefectos.Close
        Set RDefectos = Nothing
        If Err <> 0 Then
        End If
End Sub

Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = "Defecto"
End Sub

Private Sub OptDescripcion_Click()
        Lbletiqueta.Caption = "Descrip"
End Sub

Private Sub TabPuestos_Click(PreviousTab As Integer)
    If TabPuestos.Tab = 0 Then
        CmdBotones.Item(4).Enabled = True
        If CmdBotones.Item(2).Enabled = False Then
            Llena_Campos
        End If
    Else
        CmdBotones.Item(4).Enabled = False
    End If
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

Public Sub Llena_Campos()
On Error Resume Next
        Txttexto.Item(0).Text = RDefectos!Defecto
        Txttexto.Item(1).Text = RDefectos!Descrip
        Txttexto.Item(2).Text = RDefectos!Tipo
        Txttexto.Item(3).Text = RDefectos!Usuario
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        Txttexto.Item(0).Text = ""
        Txttexto.Item(1).Text = ""
        Txttexto.Item(2).Text = ""
        Txttexto.Item(3).Text = ""
End Sub
