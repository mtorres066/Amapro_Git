VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BodegasProductoTerminado 
   BackColor       =   &H00008000&
   Caption         =   "Bodegas De Inventario"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "BodegasProductoTerminado.frx":0000
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
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   3495
         Left            =   120
         TabIndex        =   41
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
         Height          =   615
         Left            =   7800
         Picture         =   "BodegasProductoTerminado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   492
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   4320
         TabIndex        =   35
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   32
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
      MouseIcon       =   "BodegasProductoTerminado.frx":24B4
      Picture         =   "BodegasProductoTerminado.frx":28F6
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Ultimo Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "BodegasProductoTerminado.frx":2E28
      Picture         =   "BodegasProductoTerminado.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Siguiente Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "BodegasProductoTerminado.frx":379C
      Picture         =   "BodegasProductoTerminado.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Registro Anterior"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "BodegasProductoTerminado.frx":4110
      Picture         =   "BodegasProductoTerminado.frx":4552
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Primer Registro"
      Top             =   3960
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   3732
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "BodegasProductoTerminado.frx":4A84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "BodegasProductoTerminado.frx":4D9E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "BodegasProductoTerminado.frx":51F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBuscar(1)"
      Tab(2).Control(1)=   "CmdBuscar(0)"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "Lbletiqueta"
      Tab(2).ControlCount=   5
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   40
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
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   732
         Index           =   1
         Left            =   -68760
         Picture         =   "BodegasProductoTerminado.frx":5642
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   1812
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   732
         Index           =   0
         Left            =   -68760
         Picture         =   "BodegasProductoTerminado.frx":594C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   1812
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   16
         ToolTipText     =   " "
         Top             =   1440
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
         TabIndex        =   23
         Top             =   960
         Width           =   2685
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   1320
            TabIndex        =   15
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos de Bodega"
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
         Height          =   2892
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   8115
         Begin VB.CheckBox Check2 
            Caption         =   "Es Bodega No Conforme"
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
            Height          =   315
            Left            =   5400
            TabIndex        =   39
            Top             =   2520
            Width           =   2532
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Es Bodega De Proceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2880
            TabIndex        =   6
            Top             =   2520
            Width           =   2412
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   5
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2160
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   6855
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
            Height          =   252
            Left            =   2880
            TabIndex        =   30
            Top             =   2160
            Width           =   5052
         End
         Begin VB.Label Label2 
            Caption         =   "Grupo"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   732
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   192
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Encargado"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Telefono"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Direccion"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   720
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
         Height          =   252
         Left            =   -70080
         TabIndex        =   24
         Top             =   1440
         Width           =   1212
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6360
      MouseIcon       =   "BodegasProductoTerminado.frx":7646
      Picture         =   "BodegasProductoTerminado.frx":7A88
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5280
      MouseIcon       =   "BodegasProductoTerminado.frx":9AFA
      Picture         =   "BodegasProductoTerminado.frx":9F3C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4200
      MouseIcon       =   "BodegasProductoTerminado.frx":A46E
      Picture         =   "BodegasProductoTerminado.frx":A8B0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3120
      MouseIcon       =   "BodegasProductoTerminado.frx":ADE2
      Picture         =   "BodegasProductoTerminado.frx":B224
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   2040
      MouseIcon       =   "BodegasProductoTerminado.frx":B756
      Picture         =   "BodegasProductoTerminado.frx":BB98
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   960
      MouseIcon       =   "BodegasProductoTerminado.frx":C0CA
      Picture         =   "BodegasProductoTerminado.frx":C50C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1000
   End
End
Attribute VB_Name = "BodegasProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEditar As Boolean
Dim VTexto As String

Dim RBodegasMateriaPrima As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
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
         CmdEditar.Enabled = True
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
        
        TxtTexto.Item(0).Enabled = True
        TxtTexto.Item(0).SetFocus
        TxtTexto.Item(6).Text = GUsuario
        BEditar = False
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RBodegasMateriaPrima.Delete
                        
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
                        RBodegasMateriaPrima.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RBodegasMateriaPrima.MoveNext
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
        RBodegasMateriaPrima.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RBodegasMateriaPrima.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RBodegasMateriaPrima.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RBodegasMateriaPrima.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RBodegasMateriaPrima.BOF Then
        RBodegasMateriaPrima.MoveFirst
    ElseIf RBodegasMateriaPrima.EOF Then
        RBodegasMateriaPrima.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    
    'INICIALIZAMOS EL RECORDSET
        Set RBodegasMateriaPrima = New ADODB.Recordset
        
    If Index = 0 Then
        If OptCodigo.Value = True Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario where CodigoBodega like '" & TxtBuscar.Text & "%'")
        ElseIf OptNombre.Value = True Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario where Descripcion like '" & TxtBuscar.Text & "%'")
        End If
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario")
    End If
        Set DataGrid1.DataSource = RBodegasMateriaPrima
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
            TxtTexto.Item(0).Enabled = True
                    
End Sub

Private Sub CmdEditar_Click()

        Bandera = True
        botones
        TxtTexto.Item(0).Enabled = False
        TxtTexto.Item(1).SetFocus
        TxtTexto.Item(6).Text = GUsuario
        BEditar = True
        
End Sub

Private Sub CmdGrabar_Click()
On Error Resume Next
                    'AGREGAR
                    If BEditar = False Then
                            VTexto = "Values('" & TxtTexto.Item(0).Text & "', '" ' CODIGO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'DESCRIPCION
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'DIRECCION
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '" 'TELEFONO
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'ENCARGADO
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', " 'GRUPO
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ES BODEGA DE PROCESO
                            Else
                                VTexto = VTexto & "0" & ", " 'ES BODEGA DE PROCESO
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ES BODEGA DE NO CONFORME
                            Else
                                VTexto = VTexto & "0" & ", '" 'ES BODEGA DE NO CONFORME
                            End If
                            
                            VTexto = VTexto & TxtTexto.Item(6).Text & "')" 'USUARIO
                            
                            Conexion.Execute "Insert Into BodegasInventario " & VTexto
                    'EDITAR
                    Else
                            VTexto = "Descripcion = '" & TxtTexto.Item(1).Text & "', " 'DESCRIPCION
                            VTexto = VTexto & "Direccion = '" & TxtTexto.Item(2).Text & "', " 'DIRECCION
                            VTexto = VTexto & "Telefono = '" & TxtTexto.Item(3).Text & "', " 'TELEFONO
                            VTexto = VTexto & "Encargado = '" & TxtTexto.Item(4).Text & "', " 'ENCARGADO
                            VTexto = VTexto & "Grupo = '" & TxtTexto.Item(5).Text & "', " 'GRUPO
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ES BODEGA DE PROCESO
                            Else
                                VTexto = VTexto & "0" & ", " 'ES BODEGA DE PROCESO
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ES BODEGA DE NO CONFORME
                            Else
                                VTexto = VTexto & "0" & ", " 'ES BODEGA DE NO CONFORME
                            End If
                            VTexto = VTexto & "usuario = '" & TxtTexto.Item(6).Text & "' " ' USUARIO
                            VTexto = VTexto & "Where CodigoBodega = '" & TxtTexto.Item(0).Text & "'"
                        
                            Conexion.Execute "UPDATE BodegasInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = 3022 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> 3022 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217900 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217900 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                        TxtTexto.Item(0).Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RBodegasMateriaPrima.Requery
   
      

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
                RBodegasMateriaPrima.Sort = RBodegasMateriaPrima.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

    
End Sub

Private Sub DBGridBusqueda_DblClick()
            TxtGrupo.Text = DBGridBusqueda.Columns(0).Text
            TxtGrupo.SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                TxtGrupo.Text = DBGridBusqueda.Columns(0).Text
                TxtGrupo.SetFocus
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
        Set RBodegasMateriaPrima = New ADODB.Recordset
        Call Abrir_Recordset(RBodegasMateriaPrima, "Select * From BodegasInventario")
        Set DataGrid1.DataSource = RBodegasMateriaPrima
        Llena_Campos
    
        If GEditar = True Then
                DataGrid1.AllowUpdate = True
        Else
                DataGrid1.AllowUpdate = False
        End If

End Sub


Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
        Lbletiqueta.Caption = "Descripcion"
End Sub



Private Sub TabBodegas_Click(PreviousTab As Integer)
    If TabBodegas.Tab = 0 Then
        If CmdGrabar.Enabled = False Then
            Llena_Campos
        End If
    End If

End Sub

Private Sub Txtbusqueda_Change()
            Set RBodegasMateriaPrimaGrupos = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    Call Abrir_Recordset(RBodegasMateriaPrimaGrupos, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos where Descripcion Like '" & TxtBusqueda.Text & "*'")
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    Call Abrir_Recordset(RBodegasMateriaPrimaGrupos, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos where Descripcion Like '*" & TxtBusqueda.Text & "*'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    Call Abrir_Recordset(RBodegasMateriaPrimaGrupos, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos where Codigo Like '" & TxtBusqueda.Text & "*'")
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    Call Abrir_Recordset(RBodegasMateriaPrimaGrupos, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos where Codigo Like '*" & TxtBusqueda.Text & "*'")
                End If
            End If
            
                    Set DBGridBusqueda.DataSource = RBodegasMateriaPrimaGrupos
                    
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next
        
        TxtTexto.Item(0).Text = RBodegasMateriaPrima!CodigoBodega
        TxtTexto.Item(1).Text = RBodegasMateriaPrima!Descripcion
        TxtTexto.Item(2).Text = RBodegasMateriaPrima!Direccion
        TxtTexto.Item(3).Text = RBodegasMateriaPrima!Telefono
        TxtTexto.Item(4).Text = RBodegasMateriaPrima!Encargado
        TxtTexto.Item(5).Text = RBodegasMateriaPrima!Grupo
        If RBodegasMateriaPrima!EsBodegaDeProceso = "-1" Then
            Check1.Value = "1"
        Else
            Check1.Value = "0"
        End If
        If RBodegasMateriaPrima!EsBodegaDeNoConforme = "-1" Then
            Check2.Value = "1"
        Else
            Check2.Value = "0"
        End If
        TxtTexto.Item(6).Text = RBodegasMateriaPrima!usuario
        
        If Err <> 0 Then
            
        End If
End Sub

Public Sub Limpia_Campos()
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(6).Text = ""
        Check1.Value = 0
        Check2.Value = 0
        
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 5 Then
            Set RBuscaGrupo = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From BodegasMateriaprimaGrupos Where Codigo = '" & TxtTexto.Item(5).Text & "'")
                If RBuscaGrupo.RecordCount > 0 Then
                    LblGrupo.Caption = RBuscaGrupo!Descripcion
                Else
                    LblGrupo.Caption = ""
                End If
        
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 5 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqued.Columns(1).Width = "4000"
                Framebuscar.Visible = True
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
            If Index = 5 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqued.Columns(1).Width = "4000"
                Framebuscar.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub
