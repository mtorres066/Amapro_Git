VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EmpleadosCapturaFaltas 
   BackColor       =   &H0080C0FF&
   Caption         =   "Captura Faltas"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "EmpleadosCapturaFaltas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   3495
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "EmpleadosCapturaFaltas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4092
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7800
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":237C
      Picture         =   "EmpleadosCapturaFaltas.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Ultimo Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":2CF0
      Picture         =   "EmpleadosCapturaFaltas.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Siguiente Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":3664
      Picture         =   "EmpleadosCapturaFaltas.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Registro Anterior"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":3FD8
      Picture         =   "EmpleadosCapturaFaltas.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Primer Registro"
      Top             =   3960
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   3732
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "EmpleadosCapturaFaltas.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "EmpleadosCapturaFaltas.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "EmpleadosCapturaFaltas.frx":50B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Empleado"
            Caption         =   "Empleado"
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
            DataField       =   "Falta"
            Caption         =   "Falta"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Observaciones"
            Caption         =   "Observacion"
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
         BeginProperty Column04 
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3420.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   975.118
            EndProperty
         EndProperty
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   8085
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   0
            Left            =   6120
            Picture         =   "EmpleadosCapturaFaltas.frx":550A
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   960
            Width           =   1812
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   1
            Left            =   6120
            Picture         =   "EmpleadosCapturaFaltas.frx":7204
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1800
            Width           =   1812
         End
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            TabIndex        =   28
            ToolTipText     =   " "
            Top             =   480
            Width           =   1845
         End
         Begin VB.Label Lbletiqueta 
            Alignment       =   1  'Right Justify
            Caption         =   "Codigo Empleado"
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
            Left            =   4200
            TabIndex        =   31
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame FrameBodegas 
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
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   8115
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1440
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   1
            ToolTipText     =   "doble click o signo + para ayuda"
            Top             =   720
            Width           =   1692
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "doble click o signo + para ayuda"
            Top             =   360
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Index           =   6
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Observacion"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label LblFal 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2880
            TabIndex        =   33
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label Label1 
            Caption         =   "Falta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   615
         End
         Begin VB.Label LblEmp 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2880
            TabIndex        =   16
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Empleado"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6240
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":750E
      Picture         =   "EmpleadosCapturaFaltas.frx":7950
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   4920
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":99C2
      Picture         =   "EmpleadosCapturaFaltas.frx":9E04
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3600
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":A336
      Picture         =   "EmpleadosCapturaFaltas.frx":A778
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   2280
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":ACAA
      Picture         =   "EmpleadosCapturaFaltas.frx":B0EC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   960
      MouseIcon       =   "EmpleadosCapturaFaltas.frx":B61E
      Picture         =   "EmpleadosCapturaFaltas.frx":BA60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1200
   End
End
Attribute VB_Name = "EmpleadosCapturaFaltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEmpleado As Boolean
Dim BFalta As Boolean
Dim VTexto As String

Dim RFaltas As New ADODB.Recordset
Dim RBuscaEmpleado As New ADODB.Recordset
Dim RBuscaFalta As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         FrameOpciones.Visible = False
         DataGrid1.Visible = False
    Else
         FrameBodegas.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        Limpia_Campos
        MskFec.Text = Date
        TxtTexto.Item(0).SetFocus
        TxtTexto.Item(6).Text = GUsuario
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RFaltas.Delete
                        
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
                        RFaltas.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RFaltas.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RFaltas.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RFaltas.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RFaltas.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RFaltas.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RFaltas.BOF Then
        RFaltas.MoveFirst
    ElseIf RFaltas.EOF Then
        RFaltas.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    Set RFaltas = New ADODB.Recordset
    If Index = 0 Then
            Call Abrir_Recordset(RFaltas, "Select * from EmpleadosCapturaFaltas where Empleado like '" & TxtBuscar.Text & "%'")
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RFaltas, "Select * from EmpleadosCapturaFaltas")
    End If
        Set DataGrid1.DataSource = RFaltas
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
                        MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
                        
                        If Not IsDate(MskFec.Text) Then
                            MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            MskFec.SetFocus
                            Exit Sub
                        End If
                        
                    Set RBuscaEmpleado = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Codigo From Empleados Where Codigo = '" & TxtTexto.Item(0).Text & "'")
                            If RBuscaEmpleado.RecordCount > 0 Then
                            
                            Else
                                MsgBox "Empleado No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(0).SetFocus
                                Exit Sub
                            End If
                     Set RBuscaFalta = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaFalta, "Select Descripcion From EmpleadosFaltas Where Codigo = '" & TxtTexto.Item(1).Text & "'")
                            If RBuscaFalta.RecordCount > 0 Then
                                
                            Else
                                MsgBox "Falta No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(1).SetFocus
                                Exit Sub
                            End If


                            VTexto = "Values('" & TxtTexto.Item(0).Text & "', '" ' CODIGO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', " 'FALTA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '"  'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', '" 'USUARIO
                            VTexto = VTexto & TxtTexto.Item(2).Text & "')" 'OBSERVACIONES
                            
                            Conexion.Execute "Insert Into EmpleadosCapturaFaltas " & VTexto
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        'IFS ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        'I ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RFaltas.Requery
                        RFaltas.MoveLast
                        Llena_Campos
      

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
                RFaltas.Sort = RFaltas.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If
    
End Sub


Private Sub DBGridBusqueda_DblClick()
        If BEmpleado = True Then
            TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(0).SetFocus
        ElseIf BFalta = True Then
            TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(1).SetFocus
        End If
            
            FrameBusqueda.Visible = False

End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
               If BEmpleado = True Then
                    TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(0).SetFocus
                ElseIf BFalta = True Then
                    TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(1).SetFocus
                End If
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
        Set RFaltas = New ADODB.Recordset
        Call Abrir_Recordset(RFaltas, "Select * From EmpleadosCapturaFaltas")
        Set DataGrid1.DataSource = RFaltas
        Llena_Campos
    
        If GEditar = True Then
                DataGrid1.AllowUpdate = True
        Else
                DataGrid1.AllowUpdate = False
        End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        RFaltas.Close
        RBuscaEmpleado.Close
        RBusqueda.Close
        RBuscaFalta.Close
        
        Set RFaltas = Nothing
        Set RBuscaEmpleado = Nothing
        Set RBusqueda = Nothing
        Set RBuscaFalta = Nothing
        
        If Err <> 0 Then
        End If
        
End Sub

Private Sub MskFec_GotFocus()
        MskFec.SelStart = 0
        MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub MskFec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub




Private Sub TabBodegas_Click(PreviousTab As Integer)
    If TabBodegas.Tab = 0 Then
        CmdBorrar.Enabled = True
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
    Else
        CmdBorrar.Enabled = False
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

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            
            If BEmpleado = True Then
                    'DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        
                    'CODIGO
                    ElseIf OptBusqueda.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            ElseIf BFalta = True Then
                    'DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosFaltas where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosFaltas where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        
                    'CODIGO
                    ElseIf OptBusqueda.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosFaltas where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosFaltas where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
                            
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next
        If RFaltas.RecordCount > 0 Then
            TxtTexto.Item(0).Text = RFaltas!Empleado
            TxtTexto.Item(1).Text = RFaltas!Falta
            MskFec.Text = RFaltas!fecha
            TxtTexto.Item(6).Text = RFaltas!Usuario
            If IsNull(RFaltas!Observaciones) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RFaltas!Observaciones
            End If
        Else
            Limpia_Campos
        End If
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        MskFec.Text = ""
        TxtTexto.Item(6).Text = ""
        TxtTexto.Item(2).Text = ""
        
End Sub

Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
        
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 0 Then
            Set RBuscaEmpleado = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtTexto.Item(0).Text & "'")
                If RBuscaEmpleado.RecordCount > 0 Then
                    LblEmp.Caption = RBuscaEmpleado!Descripcion
                Else
                    LblEmp.Caption = ""
                End If
        ElseIf Index = 1 Then
            Set RBuscaFalta = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaFalta, "Select Descripcion From EmpleadosFaltas Where Codigo = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaFalta.RecordCount > 0 Then
                    LblFal.Caption = RBuscaFalta!Descripcion
                Else
                    LblFal.Caption = ""
                End If
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        If Index = 0 Then
                BEmpleado = True
                BFalta = False
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaEmpleado = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaEmpleado, "Select Codigo, Descripcion From Empleados")
                Set DBGridBusqueda.DataSource = RBuscaEmpleado
        ElseIf Index = 1 Then
                BEmpleado = False
                BFalta = True
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaFalta = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaFalta, "Select Codigo, Descripcion From EmpleadosFaltas")
                Set DBGridBusqueda.DataSource = RBuscaFalta
        End If
        
        If Index = 0 Or Index = 1 Then
                'LLENAMOS EL GRID CON EL RECORDSET
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
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
        
        If KeyAscii = 43 Then
            If Index = 0 Then
                            BEmpleado = True
                            BFalta = False
                            'INICIALIZAMOS EL RECORDSET
                            Set RBuscaEmpleado = New ADODB.Recordset
                            'ABRIMOS EL RECORDSET
                            Call Abrir_Recordset(RBuscaEmpleado, "Select Codigo, Descripcion From Empleados")
                            Set DBGridBusqueda.DataSource = RBuscaEmpleado
                    ElseIf Index = 1 Then
                            BEmpleado = False
                            BFalta = True
                            'INICIALIZAMOS EL RECORDSET
                            Set RBuscaFalta = New ADODB.Recordset
                            'ABRIMOS EL RECORDSET
                            Call Abrir_Recordset(RBuscaFalta, "Select Codigo, Descripcion From EmpleadosFaltas")
                            Set DBGridBusqueda.DataSource = RBuscaFalta
                    End If
                    
                    If Index = 0 Or Index = 1 Then
                            'LLENAMOS EL GRID CON EL RECORDSET
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                    End If
        End If
End Sub
