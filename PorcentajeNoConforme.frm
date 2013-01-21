VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PorcentajeNoConforme 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentaje Conforme De Pedidos"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "PorcentajeNoConforme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda de Datos"
      Height          =   5895
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4695
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8281
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
      Begin VB.TextBox TxtBus 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   6735
      End
      Begin VB.OptionButton OptBus 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptBus 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   8160
         Picture         =   "PorcentajeNoConforme.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Sale de Busqueda"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblBus 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6720
      MouseIcon       =   "PorcentajeNoConforme.frx":3D6C
      Picture         =   "PorcentajeNoConforme.frx":41AE
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5640
      MouseIcon       =   "PorcentajeNoConforme.frx":6220
      Picture         =   "PorcentajeNoConforme.frx":6662
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4560
      MouseIcon       =   "PorcentajeNoConforme.frx":6B94
      Picture         =   "PorcentajeNoConforme.frx":6FD6
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3480
      MouseIcon       =   "PorcentajeNoConforme.frx":7508
      Picture         =   "PorcentajeNoConforme.frx":794A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2400
      MouseIcon       =   "PorcentajeNoConforme.frx":7E7C
      Picture         =   "PorcentajeNoConforme.frx":82BE
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1320
      MouseIcon       =   "PorcentajeNoConforme.frx":87F0
      Picture         =   "PorcentajeNoConforme.frx":8C32
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   8160
      MouseIcon       =   "PorcentajeNoConforme.frx":9164
      Picture         =   "PorcentajeNoConforme.frx":95A6
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Ultimo Registro"
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7800
      MouseIcon       =   "PorcentajeNoConforme.frx":9AD8
      Picture         =   "PorcentajeNoConforme.frx":9F1A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Siguiente Registro"
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   840
      MouseIcon       =   "PorcentajeNoConforme.frx":A44C
      Picture         =   "PorcentajeNoConforme.frx":A88E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Registro Anterior"
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   480
      MouseIcon       =   "PorcentajeNoConforme.frx":ADC0
      Picture         =   "PorcentajeNoConforme.frx":B202
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Primer Registro"
      Top             =   5040
      Width           =   495
   End
   Begin TabDlg.SSTab TabDesperdicio 
      Height          =   4815
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "PorcentajeNoConforme.frx":B734
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDesperdicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "PorcentajeNoConforme.frx":BA4E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridDesperdicio"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "PorcentajeNoConforme.frx":BEA0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGridDesperdicio 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "NumeroDocumento"
            Caption         =   "# Documento"
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
            DataField       =   "TipoDocumento"
            Caption         =   "TipoDocumento"
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
            DataField       =   "Pedido"
            Caption         =   "Pedido"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "PorcentajeConforme"
            Caption         =   "% Conforme"
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
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1170.142
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameBusquedadeDatos 
         Caption         =   "Busqueda de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Width           =   8775
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   7
            Left            =   6840
            Picture         =   "PorcentajeNoConforme.frx":C2F2
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3240
            Width           =   1812
         End
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   6
            Left            =   6840
            Picture         =   "PorcentajeNoConforme.frx":C5FC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2400
            Width           =   1812
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   7
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label LblBusqueda 
            Alignment       =   1  'Right Justify
            Caption         =   "No Documento"
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
            Left            =   5160
            TabIndex        =   13
            Top             =   2040
            Width           =   1575
         End
      End
      Begin VB.Frame FrameDesperdicio 
         Caption         =   "Datos del Desperdicio"
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
         Height          =   3255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   0
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   8
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1440
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   4
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   3
            ToolTipText     =   "Doble click o signo '+' para ayuda"
            Top             =   1680
            Width           =   1440
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   12
            TabIndex        =   2
            ToolTipText     =   "Doble click o signo '+' para ayuda"
            Top             =   1320
            Width           =   1440
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   1
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label LblLinea 
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
            Left            =   3000
            TabIndex        =   21
            Top             =   960
            Width           =   5415
         End
         Begin VB.Label LblFichaTecnica 
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
            Left            =   3000
            TabIndex        =   11
            Top             =   1680
            Width           =   5415
         End
         Begin VB.Label LblProceso 
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
            Left            =   3000
            TabIndex        =   10
            Top             =   1320
            Width           =   5415
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "% Conforme"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   20
            Top             =   2040
            Width           =   840
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   19
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "# Pedido"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   1020
            Width           =   1185
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "# Documento"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   645
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "PorcentajeNoConforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim BDocumento As Boolean
Dim BPedido As Boolean
Dim BCodigo As Boolean
Dim BEditar As Boolean

Dim RPedidos As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaPedido As New ADODB.Recordset
Dim RBuscaCodigo As New ADODB.Recordset
Dim RBuscaPedido2 As New ADODB.Recordset

Dim VUltimaLinea As String
Dim VUltimaFichaTecnica As String
Dim VUltimoTurno As String
Dim VUltimaFecha As String
Dim VTexto As String


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    
    
        'AGREGAR
        If Index = 0 Then
                Bandera = True
                botones
                Limpia_Campos
                
                'HABILITA LA LLAVE
                TxtTexto.Item(2).Enabled = True
                TxtTexto.Item(3).Enabled = True
                TxtTexto.Item(0).SetFocus
                TxtTexto.Item(8).Text = GUsuario
                BEditar = False
        'EDITAR
        ElseIf Index = 1 Then
        
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                TxtTexto.Item(2).Enabled = False
                TxtTexto.Item(3).Enabled = False
                TxtTexto.Item(0).SetFocus
                TxtTexto.Item(8).Text = GUsuario
                BEditar = True
        'GRABAR
        ElseIf Index = 2 Then
        
                    If Not IsNumeric(TxtTexto.Item(4).Text) Then
                            MsgBox "Porcentaje Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                                                
                    Set RBuscaPedido2 = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaPedido2, "Select * From DetallePedidosProveedores Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(3).Text & "'")
                            If RBuscaPedido2.RecordCount > 0 Then
                            
                            Else
                                MsgBox "Pedido Con Este Codigo, No Existe En Los Pedidos De Proveedor", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                            End If
                        
                    'AGREGAR
                    If BEditar = False Then
                            VTexto = "'" & TxtTexto.Item(0).Text & "', '" 'NUMERO DOCUMENTO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'TIPO DOCUMENTO
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'PEDIDO
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', " 'CODIGO
                            VTexto = VTexto & TxtTexto.Item(4).Text & ", '" 'PORCENTAJE CONFORME
                            VTexto = VTexto & TxtTexto.Item(8).Text & "'" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into PedidosProveedoresPorcentajeNo Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            VTexto = "NumeroDocumento = '" & TxtTexto.Item(0).Text & "', " 'NUMERO DOCUMENTO
                            VTexto = VTexto & "TipoDocumento = '" & TxtTexto.Item(1).Text & "', " 'TIPO DE DOCUMENTO
                            VTexto = VTexto & "PorcentajeConforme = " & TxtTexto.Item(4).Text & ", " '% conforme
                            VTexto = VTexto & "usuario = '" & TxtTexto.Item(8).Text & "'" ' USUARIO
                            VTexto = VTexto & " Where Pedido = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(3).Text & "'"
                        
                            Conexion.Execute "UPDATE PedidosProveedoresPorcentajeNo SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Pedido y Codigo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Pedido y Codigo Ya Existe, Ya Existe", vbOKOnly + vbInformation, "Informacion"
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
                        CmdBotones.Item(0).SetFocus
                        
                        'HABILITA LA LLAVE
                        TxtTexto.Item(2).Enabled = True
                        TxtTexto.Item(3).Enabled = True
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RPedidos.Requery
                        RPedidos.MoveLast
                        Llena_Campos

        'CANCELAR
        ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    'HABILITA LA LLAVE
                    TxtTexto.Item(2).Enabled = True
                    TxtTexto.Item(3).Enabled = True
                    
        ElseIf Index = 4 Then ' BORRAR
        
                On Error Resume Next
            VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If VMensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RPedidos.Delete
                        
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
                        RPedidos.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RPedidos.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    Set RPedidos = New ADODB.Recordset
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RPedidos, "Select * From PedidosProveedoresPorcentajeNo Where NumeroDocumento = '" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RPedidos, "Select * From PedidosProveedoresPorcentajeNo Where UPPER(NumeroDocumento) Like '" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                                Set DGridDesperdicio.DataSource = RPedidos
                                TabDesperdicio.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    Set RPedidos = New ADODB.Recordset
                    Call Abrir_Recordset(RPedidos, "Select * From PedidosProveedoresPorcentajeNo")
                    Set DGridDesperdicio.DataSource = RPedidos
                    TabDesperdicio.Tab = 1
        End If
    
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameDesperdicio.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         DGridDesperdicio.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameDesperdicio.Enabled = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         DGridDesperdicio.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RPedidos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RPedidos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RPedidos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RPedidos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RPedidos.BOF Then
        RPedidos.MoveFirst
    ElseIf RPedidos.EOF Then
        RPedidos.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DBGridBusqueda_DblClick()
    If BDocumento = True Then
        TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
        TxtTexto.Item(1).SetFocus
    ElseIf BPedido = True Then
        TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0).Text
        TxtTexto.Item(2).SetFocus
    ElseIf BCodigo = True Then
        TxtTexto.Item(3).Text = DbGridBusqueda.Columns(1).Text
        TxtTexto.Item(3).SetFocus
    End If
        FrameBusqueda.Visible = False
        
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
            If BDocumento = True Then
                TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(1).SetFocus
            ElseIf BPedido = True Then
                TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
            ElseIf BCodigo = True Then
                TxtTexto.Item(3).Text = DbGridBusqueda.Columns(1).Text
                TxtTexto.Item(3).SetFocus
            End If
                FrameBusqueda.Visible = False
    End If
End Sub

Private Sub dgriddesperdicio_HeadClick(ByVal ColIndex As Integer)
        RPedidos.Sort = RPedidos.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
    Set RPedidos = New ADODB.Recordset
        Call Abrir_Recordset(RPedidos, "Select * From PedidosProveedoresPorcentajeNo")
    Set DGridDesperdicio.DataSource = RPedidos
    Llena_Campos
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    RPedidos.Close
    RBusqueda.Close
    RBuscaDocumento.Close
    RBuscaPedido.Close
    RBuscaCodigo.Close
    
    Set RPedidos = Nothing
    Set RBusqueda = Nothing
    Set RBuscaDocumento = Nothing
    Set RBuscaPedido = Nothing
    Set RBuscaCodigo = Nothing
    
    If Err <> 0 Then
    End If

End Sub



Private Sub TabDesperdicio_Click(PreviousTab As Integer)
        If TabDesperdicio.Tab = 0 Then
            CmdBotones.Item(4).Enabled = True
                If CmdBotones.Item(2).Enabled = False Then
                    Llena_Campos
                End If
        Else
            CmdBotones.Item(4).Enabled = False
        End If
        
        
End Sub

Private Sub TxtBus_Change()
            
            
                    'OPCION POR DESCRIPCION
                    If OptBus.Item(0).Value = True Then
                                Set RBusqueda = New ADODB.Recordset
                                If BDocumento = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where Descripcion Like '%" & TxtBus.Text & "%'")
                                    Else 'oracle
                                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(Descripcion) Like '%" & UCase(TxtBus.Text) & "%'")
                                    End If
                                ElseIf BPedido = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select Pedido, Fecha from EncabezadoPedidosProveedores Where Documento Like '%" & TxtBus.Text & "%'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBusqueda, "Select Pedido, Fecha from EncabezadoPedidosProveedores Where UPPER(Documento) Like '%" & UCase(TxtBus.Text) & "%'")
                                    End If
                                ElseIf BCodigo = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Descrip Like '%" & TxtBus.Text & "%'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBus.Text) & "%'")
                                    End If
                                End If
                    'OPCION DE CODIGO
                    ElseIf OptBus.Item(1).Value = True Then
                                If BDocumento = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where CodigoDocumento Like '*" & TxtBus.Text & "*'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(CodigoDocumento) Like '*" & UCase(TxtBus.Text) & "*'")
                                    End If
                                ElseIf BPedido = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select Pedido, Fecha from EncabezadoPedidosProveedores Where Documento Like '%" & TxtBus.Text & "%'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBusqueda, "Select Pedido, Fecha from EncabezadoPedidosProveedores Where UPPER(Documento) Like '%" & UCase(TxtBus.Text) & "%'")
                                    End If
                                ElseIf BCodigo = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Esp_Tec Like '%" & TxtBus.Text & "%'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBus.Text) & "%'")
                                    End If
                                End If
                    End If
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(1).Width = "5000"

End Sub

Private Sub TxtBus_GotFocus()
        TxtBus.SelStart = 0
        TxtBus.SelLength = Len(TxtBus.Text)
End Sub

Private Sub TxtBus_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
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
        'CodigoDocumento
        If Index = 1 Then
            Set RBuscaDocumento = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaDocumento, "Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaDocumento.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaDocumento!Descripcion
                Else
                    LblLinea.Caption = ""
                End If
        'PROCESO
        ElseIf Index = 2 Then
            Set RBuscaPedido = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaPedido, "Select Fecha From EncabezadoPedidosProveedores Where Documento = '" & TxtTexto.Item(2).Text & "'")
                If RBuscaPedido.RecordCount > 0 Then
                    LblProceso.Caption = RBuscaPedido!fecha
                Else
                    LblProceso.Caption = ""
                End If
        'CODIGO
        ElseIf Index = 3 Then
            Set RBuscaCodigo = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaCodigo, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(3).Text & "'")
                If RBuscaCodigo.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaCodigo!Descrip
                Else
                    LblFichaTecnica.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    If Index = 1 Or Index = 2 Or Index = 3 Then
        Set RBusqueda = New ADODB.Recordset
    End If
    
    'CodigoDocumento
    If Index = 1 Then
        BDocumento = True
        BPedido = False
        BCodigo = False
        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos")
    'PROCESOS
    ElseIf Index = 2 Then
        BDocumento = False
        BPedido = True
        BCodigo = False
        Call Abrir_Recordset(RBusqueda, "Select Documento, Fecha From EncabezadoPedidosProveedores")
    'CODIGO
    ElseIf Index = 3 Then
        BDocumento = False
        BPedido = False
        BCodigo = True
        Call Abrir_Recordset(RBusqueda, "Select * From DetallePedidosProveedores Where Documento = '" & TxtTexto.Item(2).Text & "'")
    End If
        
    If Index = 1 Or Index = 2 Or Index = 3 Then
        Set DbGridBusqueda.DataSource = RBusqueda
        FrameBusqueda.Visible = True
        TxtBus.SetFocus
        DbGridBusqueda.Columns(1).Width = "4000"
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
                If Index = 1 Or Index = 2 Or Index = 3 Then
                    Set RBusqueda = New ADODB.Recordset
                End If
                
                'CodigoDocumento
                If Index = 1 Then
                    BDocumento = True
                    BPedido = False
                    BCodigo = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos")
                'PROCESOS
                ElseIf Index = 2 Then
                    BDocumento = False
                    BPedido = True
                    BCodigo = False
                    Call Abrir_Recordset(RBusqueda, "Select Documento, Fecha From EncabezadoPedidosProveedores")
                'CODIGO
                ElseIf Index = 3 Then
                    BDocumento = False
                    BPedido = False
                    BCodigo = True
                    Call Abrir_Recordset(RBusqueda, "Select * From DetallePedidosProveedores Where Documento = '" & TxtTexto.Item(2).Text & "'")
                End If
                    
                If Index = 1 Or Index = 2 Or Index = 3 Then
                    Set DbGridBusqueda.DataSource = RBusqueda
                    FrameBusqueda.Visible = True
                    TxtBus.SetFocus
                    DbGridBusqueda.Columns(0).Width = "1000"
                    DbGridBusqueda.Columns(1).Width = "4000"
                End If
                    
                
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        
        'NUMERO DOCUMEENTO
            If IsNull(RPedidos!NumeroDocumento) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RPedidos!NumeroDocumento
            End If
        'TIPO DE DOCUMENTO
            If IsNull(RPedidos!TipoDocumento) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RPedidos!TipoDocumento
            End If
        'PEDIDO
            If IsNull(RPedidos!Pedido) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RPedidos!Pedido
            End If
        'FICHA TECNICA
            If IsNull(RPedidos!Codigo) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RPedidos!Codigo
            End If
            
        TxtTexto.Item(4).Text = RPedidos!PorcentajeConforme
        TxtTexto.Item(8).Text = RPedidos!Usuario
        
        If Err <> 0 Then
            'MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
        
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = 0
        TxtTexto.Item(8).Text = ""
        
End Sub
