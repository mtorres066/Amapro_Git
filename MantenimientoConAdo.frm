VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MantenimientoConAdo 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ado"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "MantenimientoConAdo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DataPuestos 
      Height          =   375
      Left            =   240
      Top             =   5160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"MantenimientoConAdo.frx":628A
      OLEDBString     =   $"MantenimientoConAdo.frx":632E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   ""
      RecordSource    =   "select * from empleadosdepartamentos"
      Caption         =   "Departamentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   4215
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "MantenimientoConAdo.frx":63D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "MantenimientoConAdo.frx":66EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridPuestos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "MantenimientoConAdo.frx":6B3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBuscar"
      Tab(2).Control(1)=   "CmdBuscar(1)"
      Tab(2).Control(2)=   "CmdBuscar(0)"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "Lbletiqueta"
      Tab(2).ControlCount=   5
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   19
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "MantenimientoConAdo.frx":6F90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "MantenimientoConAdo.frx":729A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         TabIndex        =   17
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
         Caption         =   "Datos Del Departamento"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Usuario"
            DataSource      =   "DataPuestos"
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
            DataSource      =   "DataPuestos"
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
            DataSource      =   "DataPuestos"
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
            TabIndex        =   20
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Depto"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   840
         End
      End
      Begin MSDataGridLib.DataGrid DBGridPuestos 
         Bindings        =   "MantenimientoConAdo.frx":76DC
         Height          =   3375
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   "Codigo"
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
            DataField       =   "Descripcion"
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
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "MantenimientoConAdo.frx":76F6
      Picture         =   "MantenimientoConAdo.frx":7B38
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
      MouseIcon       =   "MantenimientoConAdo.frx":9BAA
      Picture         =   "MantenimientoConAdo.frx":9FEC
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
      MouseIcon       =   "MantenimientoConAdo.frx":A51E
      Picture         =   "MantenimientoConAdo.frx":A960
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
      MouseIcon       =   "MantenimientoConAdo.frx":AE92
      Picture         =   "MantenimientoConAdo.frx":B2D4
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
      MouseIcon       =   "MantenimientoConAdo.frx":B806
      Picture         =   "MantenimientoConAdo.frx":BC48
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
      MouseIcon       =   "MantenimientoConAdo.frx":C17A
      Picture         =   "MantenimientoConAdo.frx":C5BC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1200
   End
End
Attribute VB_Name = "MantenimientoConAdo"
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
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         Txtbuscar.Visible = False
         DataPuestos.Visible = False
         FrameOpciones.Visible = False
         DBGridPuestos.Visible = False
    Else
         FramePuestos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         Txtbuscar.Visible = True
         DataPuestos.Visible = True
         FrameOpciones.Visible = True
         DBGridPuestos.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataPuestos.Recordset
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
                    TxtTexto.Item(2).Text = GUsuario
            'EDITAR
            ElseIf Index = 1 Then
                    'EDITA EL REGISTRO
                    
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(0).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                   
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "Codigo De Departamento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(0).SetFocus
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
                    mensaje = MsgBox("�Est� seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminaci�n de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        DataPuestos.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataPuestos.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataPuestos.Recordset.EOF Then
                        DataPuestos.Recordset.MoveLast
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
    With DataPuestos
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                .RecordSource = ("Select * from EmpleadosDepartamentos where CodigoGrupo Like '*" & Txtbuscar.Text & "*'")
                .Refresh
                DBGridPuestos.Refresh
            ElseIf OptDescripcion.Value = True Then
                .RecordSource = ("Select * from EmpleadosDepartamentos where Descripcion Like '*" & Txtbuscar.Text & "*'")
                .Refresh
                DBGridPuestos.Refresh
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From EmpleadosDepartamentos"
                .Refresh
                DBGridPuestos.Refresh
        End If
    End With
        TabPuestos.Tab = 1
End Sub


Private Sub dbgridpuestos_HeadClick(ByVal ColIndex As Integer)
    DataPuestos.RecordSource = ("Select * from EmpleadosDepartamentos order by " & DBGridPuestos.Columns(ColIndex).DataField)
    DataPuestos.Refresh
    DBGridPuestos.Refresh
End Sub


Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        'DataPuestos.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        'DataPuestos.DatabaseName = BasedeDatos

End Sub

Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = "Codigo"
End Sub

Private Sub OptDescripcion_Click()
        Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_GotFocus()
        Txtbuscar.SelStart = 0
        Txtbuscar.SelLength = Len(Txtbuscar.Text)
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
