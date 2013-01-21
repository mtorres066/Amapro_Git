VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Paros 
   BackColor       =   &H0080C0FF&
   Caption         =   "Paros"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "Paros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8610
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
      Height          =   7695
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   6495
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11456
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
         Picture         =   "Paros.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   2160
      MouseIcon       =   "Paros.frx":237C
      Picture         =   "Paros.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "Paros.frx":2CF0
      Picture         =   "Paros.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ultimo Registro"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "Paros.frx":3664
      Picture         =   "Paros.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Siguiente Registro"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "Paros.frx":3FD8
      Picture         =   "Paros.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Registro Anterior"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "Paros.frx":494C
      Picture         =   "Paros.frx":4D8E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Primer Registro"
      Top             =   7080
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   6735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   8438015
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "Paros.frx":52C0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Paros.frx":55DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Paros.frx":5A2C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   26
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10398
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "CodigoParo"
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
            DataField       =   "DescripcionParo"
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
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Tipo2"
            Caption         =   "Tipo2"
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
            DataField       =   "Grupo"
            Caption         =   "Grupo"
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
         BeginProperty Column05 
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4034.835
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   794.835
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
         Height          =   5775
         Left            =   -74880
         TabIndex        =   16
         Top             =   840
         Width           =   8085
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   0
            Left            =   6120
            Picture         =   "Paros.frx":5E7E
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3600
            Width           =   1812
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   1
            Left            =   6120
            Picture         =   "Paros.frx":7B78
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   4440
            Width           =   1812
         End
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            TabIndex        =   31
            ToolTipText     =   " "
            Top             =   3120
            Width           =   1845
         End
         Begin VB.Label Lbletiqueta 
            Alignment       =   1  'Right Justify
            Caption         =   "Descripcion"
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
            Left            =   4800
            TabIndex        =   34
            Top             =   3120
            Width           =   1215
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos Del Paro"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   8115
         Begin VB.ComboBox CboTip2 
            Height          =   315
            ItemData        =   "Paros.frx":7E82
            Left            =   1080
            List            =   "Paros.frx":7E95
            TabIndex        =   3
            Text            =   "S"
            Top             =   1440
            Width           =   615
         End
         Begin VB.ComboBox CboTip 
            Height          =   315
            ItemData        =   "Paros.frx":7EAA
            Left            =   1080
            List            =   "Paros.frx":7EB7
            TabIndex        =   2
            Text            =   "S"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   2
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1800
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
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
            MaxLength       =   20
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo 2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label LblGrupo 
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
            Height          =   285
            Left            =   2880
            TabIndex        =   19
            Top             =   1800
            Width           =   5055
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   435
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6480
      MouseIcon       =   "Paros.frx":7EC4
      Picture         =   "Paros.frx":8306
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5400
      MouseIcon       =   "Paros.frx":A378
      Picture         =   "Paros.frx":A7BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4320
      MouseIcon       =   "Paros.frx":ACEC
      Picture         =   "Paros.frx":B12E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3240
      MouseIcon       =   "Paros.frx":B660
      Picture         =   "Paros.frx":BAA2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   1080
      MouseIcon       =   "Paros.frx":BFD4
      Picture         =   "Paros.frx":C416
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1000
   End
End
Attribute VB_Name = "Paros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEditar As Boolean
Dim vtexto As String
Dim Vllave As String
Dim VllaveNueva As String

Dim RParos As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscamaximo As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
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
         CmdEditar.Enabled = True
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


Private Sub CboTip_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CboTip2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        TabBodegas.Tab = 0
        Bandera = True
        botones
        Limpia_Campos
        
        'Set RBuscamaximo = New ADODB.Recordset
        '    Call Abrir_Recordset(RBuscamaximo, "Select Max(CodigoParo) from Paros where codigoparo <> 'P'")
            
        '    If RBuscamaximo.RecordCount > 0 Then
        '        TxtTexto.Item(0).Text = RBuscamaximo(0)
        '    End If
        TxtTexto.Item(0).Enabled = True
        TxtTexto.Item(0).SetFocus
        TxtTexto.Item(6).Text = GUsuario
        BEditar = False
        CboTip.Text = "S"
        CboTip2.Text = "S"
        
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RParos.Delete
                        
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
                        RParos.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RParos.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

End Sub


Private Sub CmdBotones_Click(Index As Integer)

End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        RParos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RParos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RParos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RParos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RParos.BOF Then
        RParos.MoveFirst
    ElseIf RParos.EOF Then
        RParos.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    
    'INICIALIZAMOS EL RECORDSET
        Set RParos = New ADODB.Recordset
        
    If Index = 0 Then
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RParos, "Select * from Paros where Descripcionparo like '%" & TxtBuscar.Text & "%'")
        Else
            Call Abrir_Recordset(RParos, "Select * from Paros where Upper(Descripcionparo) like '%" & UCase(TxtBuscar.Text) & "%'")
        End If
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RParos, "Select * from Paros")
    End If
        Set DataGrid1.DataSource = RParos
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
            'HABILITA LA LLAVE
            TxtTexto.Item(0).Enabled = True
End Sub

Private Sub CmdEditar_Click()
                TabBodegas.Tab = 0
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = False
                TxtTexto.Item(1).SetFocus
                TxtTexto.Item(6).Text = GUsuario
                BEditar = True
                Vllave = TxtTexto.Item(0).Text

End Sub

Private Sub CmdGrabar_Click()
On Error Resume Next
                    VllaveNueva = TxtTexto.Item(0).Text
                    
                    If CboTip.Text <> "S" And CboTip.Text <> "N" And CboTip.Text <> "P" Then
                        MsgBox "Tipo De Paro Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                    
                    If CboTip.Text <> "S" And CboTip.Text <> "N" And CboTip.Text <> "P" And CboTip.Text <> "CF" And CboTip.Text <> "MP" Then
                        MsgBox "Tipo 2 De Paro Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                    
                    Set RBuscaGrupo = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ParosGrupos Where CodigoGrupo = '" & TxtTexto.Item(2).Text & "'")
                            If RBuscaGrupo.RecordCount > 0 Then
                            
                            Else
                                MsgBox "Grupo No Existe", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                            End If
                    
                    
                    If BEditar = False Then
                    
                            Set RBuscaGrupo = New ADODB.Recordset
                                Call Abrir_Recordset(RBuscaGrupo, "Select CodigoParo From Paros Where Codigoparo = '" & TxtTexto.Item(0).Text & "'")
                                    If RBuscaGrupo.RecordCount > 0 Then
                                        MsgBox "Codigo De Paro Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                        Exit Sub
                                    End If
                            
                            vtexto = "'" & TxtTexto.Item(0).Text & "', " ' CODIGO
                            vtexto = vtexto & "'" & TxtTexto.Item(1).Text & "', " ' DESCRIPCION
                            vtexto = vtexto & "'" & CboTip.Text & "', " ' TIPO
                            vtexto = vtexto & "'" & TxtTexto.Item(6).Text & "', " ' USUARIO
                            vtexto = vtexto & "'" & TxtTexto.Item(2).Text & "', " 'GRUPO
                            vtexto = vtexto & "'" & CboTip2.Text & "'" 'TIPO 2
                            
                            Conexion.Execute "Insert Into Paros Values(" & vtexto & ")"
                    Else
                            vtexto = "CodigoParo = '" & TxtTexto.Item(0).Text & "', " ' CODIGO
                            vtexto = vtexto & "DescripcionParo = '" & TxtTexto.Item(1).Text & "', " ' DESCRIPCION
                            vtexto = vtexto & "Tipo = '" & CboTip.Text & "', " ' TIPO
                            vtexto = vtexto & "Usuario = '" & TxtTexto.Item(6).Text & "', " ' USUARIO
                            vtexto = vtexto & "Grupo = '" & TxtTexto.Item(2).Text & "', " 'GRUPO
                            vtexto = vtexto & "Tipo2 = '" & CboTip2.Text & "'" 'TIPO 2
                            vtexto = vtexto & " Where CodigoParo = '" & Vllave & "'" 'LLAVE
                            
                            Conexion.Execute "Update Paros Set " & vtexto
                    End If
                    
                   'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Error " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo Paro Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                 
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                        'HABILITA LA LLAVE
                        TxtTexto.Item(0).Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RParos.Requery
                        If BEditar = True Then
                            RParos.Find "CodigoParo = '" & Vllave & "'"
                        Else
                            RParos.Find "CodigoParo = '" & VllaveNueva & "'"
                        End If
                        
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
                RParos.Sort = RParos.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

    
End Sub

Private Sub DBGridBusqueda_DblClick()
            TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(2).SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
        Set RParos = New ADODB.Recordset
        Call Abrir_Recordset(RParos, "Select * From Paros")
        Set DataGrid1.DataSource = RParos
        Llena_Campos
    
        If GEditar = True Then
                DataGrid1.AllowUpdate = True
        Else
                DataGrid1.AllowUpdate = False
        End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        RParos.Close
        RBuscaGrupo.Close
        RBusqueda.Close
        
        Set RParos = Nothing
        Set RBuscaGrupo = Nothing
        Set RBusqueda = Nothing
        
        If Err <> 0 Then
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
            Set RParos = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "Select CodigoParo, DescripcionParo From Paros where DescripcionParo Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "Select CodigoParo, DescripcionParo From Paros where UPPER(DescripcionParo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "Select CodigoParo, DescripcionParo From Paros where CodigoParo Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "Select CodigoParo, DescripcionParo From Paros where UPPER(CodigoParo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
                    
                    Set DBGridBusqueda.DataSource = RParos
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next
        
        TxtTexto.Item(0).Text = RParos!CodigoParo
        TxtTexto.Item(1).Text = RParos!DescripcionParo
        CboTip.Text = RParos!Tipo
        TxtTexto.Item(6).Text = RParos!Usuario
        TxtTexto.Item(2).Text = RParos!Grupo
        CboTip2.Text = RParos!Tipo2
            
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(6).Text = ""
        
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
        If Index = 2 Then
            Set RBuscaGrupo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ParosGrupos Where CodigoGrupo = '" & TxtTexto.Item(2).Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ParosGrupos Where UPPER(CodigoGrupo) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                End If
                If RBuscaGrupo.RecordCount > 0 Then
                    LblGrupo.Caption = RBuscaGrupo!Descripcion
                Else
                    LblGrupo.Caption = ""
                End If
        
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        If Index = 2 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select CodigoGrupo, Descripcion From ParosGrupos")
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
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
            If Index = 2 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select CodigoGrupo, Descripcion From ParosGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub
