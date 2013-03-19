VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentasMetas 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meta De Ventas Por Mes"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "VentasMetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8700
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
      Height          =   6015
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   4815
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8493
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
         Picture         =   "VentasMetas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   7215
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "VentasMetas.frx":237C
      Picture         =   "VentasMetas.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Ultimo Registro"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "VentasMetas.frx":2CF0
      Picture         =   "VentasMetas.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Siguiente Registro"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "VentasMetas.frx":3664
      Picture         =   "VentasMetas.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Registro Anterior"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "VentasMetas.frx":3FD8
      Picture         =   "VentasMetas.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Primer Registro"
      Top             =   5400
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   -2147483645
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "VentasMetas.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "VentasMetas.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "VentasMetas.frx":50B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   25
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   7435
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
         Height          =   4215
         Left            =   -74880
         TabIndex        =   16
         Top             =   720
         Width           =   8085
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   6120
            TabIndex        =   34
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   71565315
            CurrentDate     =   38385
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Ficha Tecnica"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Bodega"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   0
            Left            =   6120
            Picture         =   "VentasMetas.frx":550A
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2520
            Width           =   1812
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   1
            Left            =   6120
            Picture         =   "VentasMetas.frx":7204
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3360
            Width           =   1812
         End
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            TabIndex        =   36
            ToolTipText     =   " "
            Top             =   1920
            Visible         =   0   'False
            Width           =   1845
         End
         Begin MSComCtl2.DTPicker DTPFecFin 
            Height          =   255
            Left            =   6120
            TabIndex        =   35
            Top             =   1560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   71565315
            CurrentDate     =   38385
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Final"
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
            TabIndex        =   40
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Inicial"
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
            TabIndex        =   39
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Lbletiqueta 
            Alignment       =   1  'Right Justify
            Caption         =   "Bodega"
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
            Left            =   3960
            TabIndex        =   30
            Top             =   1920
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos De La Meta"
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
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   8115
         Begin MSMask.MaskEdBox MskMetTon 
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   2160
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCli 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1080
            Width           =   1692
         End
         Begin MSComCtl2.DTPicker DtpMes 
            Height          =   255
            Left            =   1080
            TabIndex        =   0
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   71565315
            CurrentDate     =   40368
         End
         Begin VB.TextBox TxtTipFicTec 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   1
            Top             =   720
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
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1692
         End
         Begin MSMask.MaskEdBox MskMetDol 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMetCan 
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Toneladas"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label LblCli 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   42
            Top             =   1080
            Width           =   5055
         End
         Begin VB.Label Label1 
            Caption         =   "Dolares"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Unidades"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label LblEmp 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Ficha Tecnica"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6360
      MouseIcon       =   "VentasMetas.frx":750E
      Picture         =   "VentasMetas.frx":7950
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1065
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5040
      MouseIcon       =   "VentasMetas.frx":7E6B
      Picture         =   "VentasMetas.frx":82AD
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3720
      MouseIcon       =   "VentasMetas.frx":8875
      Picture         =   "VentasMetas.frx":8CB7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   2400
      MouseIcon       =   "VentasMetas.frx":91EE
      Picture         =   "VentasMetas.frx":9630
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   1080
      MouseIcon       =   "VentasMetas.frx":9B8C
      Picture         =   "VentasMetas.frx":9FCE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1200
   End
End
Attribute VB_Name = "VentasMetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BFichaTecnica As Boolean
Dim BBodega As Boolean
Dim BCliente As Boolean

Dim VTexto As String

Dim RInventario As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset

Dim VUltimaFicha As String
Dim VFecha As Date
Dim VUltimoCliente As String

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtTipFicTec.SetFocus
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
        TabBodegas.Tab = 0
        Bandera = True
        botones
        Limpia_Campos
        
        TxtTipFicTec.SetFocus
        TxtTexto.Item(6).Text = GUsuario
        DtpMes.Value = Date
        TxtCli.Text = VUltimoCliente
        
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RInventario.Delete
                        
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
                        RInventario.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RInventario.MoveLast
                        
                        Llena_Campos
                    End If

End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RInventario.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RInventario.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RInventario.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RInventario.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RInventario.BOF Then
        RInventario.MoveFirst
    ElseIf RInventario.EOF Then
        RInventario.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    Set RInventario = New ADODB.Recordset
    If Index = 0 Then
            If OptOpcion.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RInventario, "Select * from VentasMetasMensuales where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And TipoFichaTecnica Like '" & TxtBuscar.Text & "%'")
                Else
                    Call Abrir_Recordset(RInventario, "Select * from VentasMetasMensuales where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(TipoFichaTecnica) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptOpcion.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RInventario, "Select * from VentasMetasMensuales where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And TipoFichaTecnica Like '" & TxtBuscar.Text & "%'")
                Else
                    Call Abrir_Recordset(RInventario, "Select * from VentasMetasMensuales where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(TipoFichaTecnica) Like " & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RInventario, "Select * from VentasMetasMensuales")
    End If
        Set DataGrid1.DataSource = RInventario
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
            
                    
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

                    VUltimoCliente = TxtCli.Text
                    
                        
                    VFecha = "01/" & Month(DtpMes.Value) & "/" & Year(DtpMes.Value)
                        
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select Codigo From FichaTecnicaTiposVentas Where Codigo = '" & TxtTipFicTec.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select Codigo From FichaTecnicaTiposVentas Where UPPER(Codigo) = '" & UCase(TxtTipFicTec.Text) & "'")
                        End If
                            If RBuscaFicha.RecordCount > 0 Then
                            
                            Else
                                MsgBox "Tipo De Ficha Tecnica No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTipFicTec.SetFocus
                                Exit Sub
                            End If
                            
                    Set RBuscaCliente = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
                        End If
                        If RBuscaCliente.RecordCount > 0 Then
                            
                        Else
                                MsgBox "Cliente No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtCli.SetFocus
                                Exit Sub
                        End If
                            
                    'REVISA QUE NO SE DUPLICA MES, TIPO, Clientes
                    Set RBuscaFicha = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFicha, "Select Fecha From VentasMetasMensuales Where Fecha = #" & Format(VFecha, "MM/dd/yyyy") & "# And TipoFichaTecnica = '" & TxtTipFicTec.Text & "' And Cliente = '" & TxtCli.Text & "'")
                            If RBuscaFicha.RecordCount > 0 Then
                                MsgBox "Mes y Tipo De Ficha Tecnica Y Cliente Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTipFicTec.SetFocus
                                Exit Sub
                            End If
                            
                            

                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(VFecha, "mm/dd/yyyy") & "#, '"  'FECHA
                            Else 'ORACLE
                                 VTexto = "To_Date('" & VFecha & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtTipFicTec.Text & "', '" ' TIPO FICHA TECNICA
                            VTexto = VTexto & TxtCli.Text & "', " ' Clientes
                            VTexto = VTexto & MskMetDol.Text & ", " 'DOLARES
                            VTexto = VTexto & MskMetCan.Text & ", " 'CANTIDAD
                            VTexto = VTexto & MskMetTon.Text & ", '"    'TONELADAS  ---------------------------------
                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                            
                            Conexion.Execute "Insert Into VentasMetasMensuales Values(" & VTexto & ")"
                    
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
                        RInventario.Requery
                        RInventario.MoveLast
                        Llena_Campos
      

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
        Llena_Campos
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
                RInventario.Sort = RInventario.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

    
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        Llena_Campos
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
        Llena_Campos
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BFichaTecnica = True Then
            TxtTipFicTec.Text = DBGridBusqueda.Columns(0).Text
            TxtTipFicTec.SetFocus
        ElseIf BCliente = True Then
            TxtCli.Text = DBGridBusqueda.Columns(0).Text
            TxtCli.SetFocus
        End If
            
            FrameBusqueda.Visible = False

End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
               If BFichaTecnica = True Then
                    TxtTipFicTec.Text = DBGridBusqueda.Columns(0).Text
                    TxtTipFicTec.SetFocus
                ElseIf BCliente = True Then
                    TxtCli.Text = DBGridBusqueda.Columns(0).Text
                    TxtCli.SetFocus
                End If
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub DtpMes_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Form_Load()
        Set RInventario = New ADODB.Recordset
        Call Abrir_Recordset(RInventario, "Select * From VentasMetasMensuales ORDER BY Fecha, TipoFichaTecnica")
        RInventario.MoveLast
        
        Set DataGrid1.DataSource = RInventario
        Llena_Campos
    
        If GEditar = True Then
                DataGrid1.AllowUpdate = True
        Else
                DataGrid1.AllowUpdate = False
        End If
        
'        DtpFecIni.Value = Date
'        DTPFecFin.Value = Date
        
        DtpFecIni.Value = RInventario!fecha
        DTPFecFin.Value = RInventario!fecha
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        RInventario.Close
        RBuscaFicha.Close
        RBusqueda.Close
        RBuscaBodega.Close
        
        Set RInventario = Nothing
        Set RBuscaFicha = Nothing
        Set RBusqueda = Nothing
        Set RBuscaBodega = Nothing
        
        If Err <> 0 Then
        End If
        
End Sub


Private Sub MskMetCan_GotFocus()
        MskMetCan.SelStart = 0
        MskMetCan.SelLength = Len(MskMetCan.Text)
End Sub

Private Sub MskMetCan_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskMetDol_GotFocus()
        MskMetDol.SelStart = 0
        MskMetDol.SelLength = Len(MskMetDol.Text)
End Sub

Private Sub MskMetDol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskMetTon_GotFocus()
        MskMetTon.SelStart = 0
        MskMetTon.SelLength = Len(MskMetTon.Text)
End Sub

Private Sub MskMetTon_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            Lbletiqueta.Caption = "Bodega"
        ElseIf Index = 1 Then
            Lbletiqueta.Caption = "Ficha Tecncia"
        End If
        TxtBuscar.SetFocus
End Sub

Private Sub TabBodegas_Click(PreviousTab As Integer)
    If TabBodegas.Tab = 0 Then
        If CmdGrabar.Enabled = False Then
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
        End If
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
            
            
                    'DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If BFichaTecnica = True Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicaTiposVentas where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            ElseIf BCliente = True Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            End If
                    'CODIGO
                    ElseIf OptBusqueda.Item(1).Value = True Then
                            If BFichaTecnica = True Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicaTiposVentas where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            ElseIf BCliente = True Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where CodigoCliente Like '%" & TxtBusqueda.Text & "%'")
                            End If
                            
                    End If
            
                            
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next
        DtpMes.Value = RInventario!fecha
        TxtTipFicTec.Text = RInventario!TipoFichaTecnica
        TxtCli.Text = RInventario!Cliente
        MskMetDol.Text = RInventario!MetaDolares
        MskMetCan.Text = RInventario!MetaCantidad
        TxtTexto.Item(6).Text = RInventario!Usuario
        MskMetTon.Text = RInventario!MetaToneladas
            
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        DtpMes.Value = Date
        TxtTipFicTec.Text = ""
        TxtCli.Text = ""
        MskMetDol.Text = "0"
        MskMetCan.Text = "0"
        TxtTexto.Item(6).Text = ""
        MskMetTon.Text = "0"
        
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

Private Sub TxtCli_Change()
                Set RBuscaCliente = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
                End If
                If RBuscaCliente.RecordCount > 0 Then
                    LblCli.Caption = RBuscaCliente!Descripcion
                Else
                    LblCli.Caption = ""
                End If
End Sub

Private Sub TxtCli_DblClick()
                BFichaTecnica = False
                BBodega = False
                BCliente = True
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaFicha = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaFicha, "Select CodigoCliente, Descripcion From Clientes")
                Set DBGridBusqueda.DataSource = RBuscaFicha
                'LLENAMOS EL GRID CON EL RECORDSET
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    
                            BFichaTecnica = False
                            BBodega = False
                            BCliente = True
                            'INICIALIZAMOS EL RECORDSET
                            Set RBuscaFicha = New ADODB.Recordset
                            'ABRIMOS EL RECORDSET
                            Call Abrir_Recordset(RBuscaFicha, "Select CodigoCliente, Descripcion From Clientes")
                            Set DBGridBusqueda.DataSource = RBuscaFicha
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



Private Sub TxtTipFicTec_Change()
                Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTiposVentas Where Codigo = '" & TxtTipFicTec.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFicha, "Select Descripcion From FichaTecnicaTiposVentas Where UPPER(Codigo) = '" & UCase(TxtTipFicTec.Text) & "'")
                End If
                If RBuscaFicha.RecordCount > 0 Then
                    LblEmp.Caption = RBuscaFicha!Descripcion
                Else
                    LblEmp.Caption = ""
                End If
                
End Sub

Private Sub TxtTipFicTec_DblClick()
                BFichaTecnica = True
                BBodega = False
                BCliente = False
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaFicha = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaFicha, "Select Codigo, Descripcion From FichaTecnicaTiposVentas order by codigo")
                Set DBGridBusqueda.DataSource = RBuscaFicha
                'LLENAMOS EL GRID CON EL RECORDSET
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
End Sub

Private Sub TxtTipFicTec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    
                            BFichaTecnica = True
                            BBodega = False
                            BCliente = False
                            'INICIALIZAMOS EL RECORDSET
                            Set RBuscaFicha = New ADODB.Recordset
                            'ABRIMOS EL RECORDSET
                            Call Abrir_Recordset(RBuscaFicha, "Select Codigo, Descripcion From FichaTecnicaTiposVentas order by codigo")
                            Set DBGridBusqueda.DataSource = RBuscaFicha
                            'LLENAMOS EL GRID CON EL RECORDSET
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                    
        End If
End Sub
