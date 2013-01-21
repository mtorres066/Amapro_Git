VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CerrarPedidoMateriaPrimaClientes 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrar Pedidos De Materia Prima De Clientes"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "CerrarPedidoMateriaPrimaClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda De Datos"
      Height          =   6135
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   8160
         Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "sale de busqueda"
         Top             =   360
         Width           =   855
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "CerrarPedidoMateriaPrimaClientes.frx":293C
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "CerrarPedidoMateriaPrimaClientes.frx":2957
         TabIndex        =   29
         ToolTipText     =   "doble click o signo '+' para seleccionar"
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Data DataCerrarPedidos 
      Caption         =   "Cerrar Pedido De Materia Prima De Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CerrarPedidoClientes"
      Top             =   5040
      Width           =   9135
   End
   Begin TabDlg.SSTab TabDepartamentos 
      Height          =   4815
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CerrarPedidoMateriaPrimaClientes.frx":3331
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCerrarPedidos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CerrarPedidoMateriaPrimaClientes.frx":364B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridCerrarPedidos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "CerrarPedidoMateriaPrimaClientes.frx":3A9D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridCerrarPedidos 
         Bindings        =   "CerrarPedidoMateriaPrimaClientes.frx":3EEF
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "CerrarPedidoMateriaPrimaClientes.frx":3F0F
         TabIndex        =   13
         ToolTipText     =   "click en encabezado columna para indexar"
         Top             =   720
         Width           =   8895
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   14
         Top             =   720
         Width           =   8775
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "No. Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   3000
            Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":518D
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   7200
            TabIndex        =   34
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24510467
            CurrentDate     =   37501
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   7200
            TabIndex        =   33
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24510467
            CurrentDate     =   37501
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   17
            Top             =   2160
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   1560
            Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":5497
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "No. Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   120
            Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":8311
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label LblFecIni 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   6000
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label LblFecFin 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   6000
            TabIndex        =   35
            Top             =   1560
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label LblBusqueda 
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
            Left            =   3840
            TabIndex        =   25
            Top             =   2160
            Width           =   2055
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   6000
            TabIndex        =   19
            Top             =   3240
            Width           =   2535
            Caption         =   "Seleccionar Todos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   6
            Left            =   6000
            TabIndex        =   18
            Top             =   2520
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":D913
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameCerrarPedidos 
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
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   8655
         Begin MSMask.MaskEdBox MskSaldoPedido 
            Height          =   285
            Left            =   3840
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FechaOperacion"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   3
            Left            =   5040
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   960
         End
         Begin MSMask.MaskEdBox MskCan 
            DataField       =   "Cantidad"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Left            =   2280
            TabIndex        =   7
            Top             =   1800
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   49152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecRec 
            DataField       =   "FechaRecepcion"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Left            =   2280
            TabIndex        =   4
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   4
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Pedido"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   2
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   6
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1440
            Width           =   1500
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CodigoMateriaPrima"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   5
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1095
            Width           =   1500
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   1
            Top             =   360
            Width           =   1500
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3840
            Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":DD65
            Top             =   480
            Width           =   480
         End
         Begin VB.Label LblSaldo 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3840
            TabIndex        =   31
            Top             =   1560
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label LblMateriaPrima 
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
            Left            =   3840
            TabIndex        =   27
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Despacho"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   26
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Pedido"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   24
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Entregada"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   23
            Top             =   1800
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Materia Prima"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   22
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "No. Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   21
            Top             =   360
            Width           =   1230
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1800
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "3175;1085"
      Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":E62F
      Accelerator     =   65
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   5640
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "3175;1085"
      Accelerator     =   71
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   3
      Left            =   3720
      TabIndex        =   10
      Top             =   5640
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "3175;1085"
      Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":EB71
      Accelerator     =   67
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   4
      Left            =   5520
      TabIndex        =   11
      Top             =   5640
      Width           =   1800
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "3175;1085"
      Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":F0B3
      Accelerator     =   66
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   5
      Left            =   7320
      TabIndex        =   12
      Top             =   5640
      Width           =   1920
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "3387;1085"
      Picture         =   "CerrarPedidoMateriaPrimaClientes.frx":F5F5
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "CerrarPedidoMateriaPrimaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim VFechaEntrada As Date
Dim VDiasDeAtraso As Long
Dim VNumeroPedido As String
Dim VCodigo As String
Dim VCantidadSalida As Long
Dim VCantidadCerrarPedido As Long

Dim RBuscaPedido As Recordset
Dim RBuscaCantidadPedido As Recordset
Dim RBuscaPedido2 As Recordset
Dim RBuscaPedido3 As Recordset
Dim RBuscaDespacho As Recordset
Dim RBuscaCantidadRecepcion As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaSaldo As Recordset

Dim BMateriaPrima As Boolean
Dim BPedido As Boolean



Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataCerrarPedidos.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                Txttexto.Item(4).Text = GUsuario
                Txttexto.Item(3).Text = Date
                Txttexto.Item(0).SetFocus
                
        'GRABAR
        ElseIf Index = 2 Then
        
                'REVISA SI ES NUMERICO LA CANTIDAD DE ENTRADA
                If Not IsNumeric(MskCan.Text) Then
                        MsgBox "Cantidad  De Entrada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                
                'REVISA SI ES FECHA VALIDA
                If Not IsDate(MskFecRec.Text) Then
                        MsgBox "Fecha De Recepcion Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                
                
                'REVISA SI EXISTE EL PEDIDO
                Set RBuscaPedido2 = Db.OpenRecordset("Select * From PedidosMateriaPrimaClientes Where Documento = '" & Txttexto.Item(2).Text & "'")
                If RBuscaPedido2.RecordCount > 0 Then
                Else
                         MsgBox "Pedido No Existe", vbOKOnly + vbInformation, "Informacion"
                End If
                
                'REVISA SI EXISTE EL PEDIDO Y CON LA MATERIA PRIMA
                Set RBuscaPedido3 = Db.OpenRecordset("Select * From PedidosMateriaPrimaClientes Where Documento = '" & Txttexto.Item(2).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                If RBuscaPedido3.RecordCount > 0 Then
                Else
                         MsgBox "Este Pedido No Corresponde A Esta Materia Prima", vbOKOnly + vbInformation, "Informacion"
                End If
                
                'BUSCA EL SALDO DEL PEDIDO
                Set RBuscaSaldo = Db.OpenRecordset("Select SaldoPorEntregar From PedidosMateriaPrimaClientes Where Documento = '" & Txttexto.Item(2).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                If RBuscaSaldo.RecordCount > 0 Then
                    If Val(MskCan.Text) > Val(RBuscaSaldo!SaldoPorEntregar) Then
                        MsgBox "La Cantidad No Puede Ser Mayor Al Saldo Del Pedido", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                End If
                
                'BUSCA SI EL CODIGO PERTENECE AL DESPACHO
                Set RBuscaDespacho = Db.OpenRecordset("Select * From DetalleEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                    If RBuscaDespacho.RecordCount > 0 Then
                    Else
                        MsgBox "El Codigo Materia Prima No Se Capturo Por Este Despacho " & Txttexto.Item(0).Text, vbOKOnly + vbExclamation, "Advertencia"
                        'Txttexto.Item(1).SetFocus
                        'Exit Sub
                    End If
                                    
                'REVISA SI LA CANTIDAD QUE ESTAN SALIENDO NO ES MAYOR QUE LE SALDO QUE LE FALTA AL DESPACHO DEPENDIENDO DEL CODIGO
                '________________________________________________________________________________________________________
                            'BUSCA LA CANTIDAD QUE TRAE EL DESPACHO
                            Set RBuscaDespacho = Db.OpenRecordset("Select Sum(Cantidad) From DetalleEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                                If RBuscaDespacho.RecordCount > 0 Then
                                    If IsNull(RBuscaDespacho(0)) Then
                                        VCantidadSalida = 0
                                    Else
                                        VCantidadSalida = RBuscaDespacho(0)
                                    End If
                                Else
                                    VCantidadSalida = 0
                                End If
                            
                            'BUSCA LA CANTIDAD QUE TIENE EL CIERRE DE PEDIDOS
                            Set RBuscaPedido = Db.OpenRecordset("Select Sum(Cantidad) From CerrarPedidoMateriaPrimaClientes Where Documento = '" & Txttexto.Item(0).Text & "' And CodigoMateriaPrima = '" & Txttexto.Item(1).Text & "'")
                                If RBuscaPedido.RecordCount > 0 Then
                                    If IsNull(RBuscaPedido(0)) Then
                                        VCantidadCerrarPedido = 0
                                    Else
                                        VCantidadCerrarPedido = RBuscaPedido(0)
                                    End If
                                Else
                                    VCantidadCerrarPedido = 0
                                End If
                            'SI LA CANTIDAD INGRESADA ES MAYOR AL SALDO DE LA CANTIDAD DE RECEPCION MENOS LO INGRESADO EN EL CIERRE DE PEDIDOS
                            If Val(MskCan.Text) > (Val(VCantidadSalida) - Val(VCantidadCerrarPedido)) Then
                                MsgBox "La Cantidad Es Mayor Que El Saldo Que Se Despacho, Se Va A Grabar El Registro Pero Verifique", vbExclamation, "Advertencia"
                                'MskCan.SetFocus
                                'Exit Sub
                            End If
                '________________________________________________________________________________________________________
                
                VCantidadSalida = MskCan.Text
                VFechaEntrada = MskFecRec.Text
                VNumeroPedido = Txttexto.Item(2).Text
                VCodigo = Txttexto.Item(1).Text
                
                'GRABA DATOS
                .Update
                
                        If Err = 3022 Then
                                MsgBox "El Codigo De Materia Prima Ya Esta Ingresado Con Este Despacho Para Este Pedido", vbOKOnly + vbExclamation, "Verifique"
                                Exit Sub
                        ElseIf Err.Number <> 0 And Err <> 3022 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                
                '---------- PEDIDO ------------------------------------------------------
                'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                Set RBuscaPedido = Db.OpenRecordset("Select CantidadEntregada, SaldoPorEntregar, FechaParaEntregar, FechaEntregaTotal, DiasDeAtraso From PedidosMateriaPrimaClientes Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                    If RBuscaPedido.RecordCount > 0 Then
                        'EDITA EL REGISTRO DE PEDIDO Y ACTUALIZA DATOS
                        RBuscaPedido.Edit
                            RBuscaPedido!CantidadEntregada = Val(RBuscaPedido!CantidadEntregada) + Val(VCantidadSalida)
                            RBuscaPedido!SaldoPorEntregar = Val(RBuscaPedido!SaldoPorEntregar) - Val(VCantidadSalida)
                                            
                            'SI EL SALDO POR ENTREGAR YA ESTA EN CERO O MENOR QUE CERO ACTUALIZA LA FECHA DE ENTREGA Y CALCULA
                            'LOS DIAS DE ATRASO
                            If RBuscaPedido!SaldoPorEntregar <= 0 Then
                                'CAMBIA LA FECHA DE ENTREGA TOTAL POR LA ACTUAL DEL ULTIMO INGRESO
                                RBuscaPedido!FechaEntregaTotal = VFechaEntrada
                                            
                                'CALCULA LOS DIAS DE ATRASO
                                VDiasDeAtraso = (DateValue(RBuscaPedido!FechaParaEntregar) - DateValue(VFechaEntrada))
                                                
                                'SI LA VARIABLE VDIASDEATRASO ES MENOR QUE CERO ES PORQUE ENTREGO EL PEDIDO ANTES DE LA FECHA
                                If VDiasDeAtraso < 0 Then
                                    VDiasDeAtraso = VDiasDeAtraso * -1
                                Else
                                    VDiasDeAtraso = 0
                                End If
                                                
                                'MODIFICA LOS DIAS DE ATRASO EN EL PEDIDO
                                RBuscaPedido!DiasDeAtraso = VDiasDeAtraso
                            Else
                                If IsNull(RBuscaPedido!FechaEntregaTotal) Then
                                Else
                                    RBuscaPedido!FechaEntregaTotal = ""
                                    RBuscaPedido!DiasDeAtraso = "0"
                                End If
                            End If
                        'GRABA DATOS
                        RBuscaPedido.Update
                    End If
                    
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
                
        'CANCELAR
        ElseIf Index = 3 Then
                .CancelUpdate
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
        
                If GBorrar = False Then
                      MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                      Exit Sub
                End If
                
                VCantidadSalida = MskCan.Text
                VFechaEntrada = MskFecRec.Text
                VNumeroPedido = Txttexto.Item(2).Text
                VCodigo = Txttexto.Item(1).Text
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If vbYes Then
                    .Delete
                    .MoveNext
                    
                            '---------- PEDIDO ------------------------------------------------------
                            'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                            Set RBuscaPedido = Db.OpenRecordset("Select CantidadEntregada, SaldoPorEntregar, FechaParaEntregar, FechaEntregaTotal, DiasDeAtraso From PedidosMateriaPrimaClientes Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                                If RBuscaPedido.RecordCount > 0 Then
                                    'EDITA EL REGISTRO DE PEDIDO Y ACTUALIZA DATOS
                                    RBuscaPedido.Edit
                                        RBuscaPedido!CantidadEntregada = RBuscaPedido!CantidadEntregada - VCantidadSalida
                                        RBuscaPedido!SaldoPorEntregar = RBuscaPedido!SaldoPorEntregar + VCantidadSalida
                                                        
                                        'SI EL SALDO POR ENTREGAR ES MAYOR QUE CERO CAMBIA LA FECHA
                                        If RBuscaPedido!SaldoPorEntregar > 0 Then
                                            'CAMBIA LA FECHA DE ENTREGA TOTAL
                                            RBuscaPedido!FechaEntregaTotal = "01/01/00"
                                            'MODIFICA LOS DIAS DE ATRASO EN EL PEDIDO
                                            RBuscaPedido!DiasDeAtraso = 0
                                        End If
                                    'GRABA DATOS
                                    RBuscaPedido.Update
                                End If
                    
                            If Err.Number = 3021 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    'DESPACHO
                    If OptBusqueda.Item(0).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoMateriaPrimaClientes where Documento = '" & TxtBusqueda.Text & "' Order By FechaRecepcion")
                    'FECHAS
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoMateriaPrimaClientes where FechaRecepcion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaRecepcion <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Order By FechaRecepcion")
                    'PEDIDO
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoMateriaPrimaClientes where Pedido = '" & TxtBusqueda.Text & "' Order By FechaRecepcion")
                    End If
                    DataCerrarPedidos.Refresh
                    DGridCerrarPedidos.Refresh
                    TabDepartamentos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataCerrarPedidos.RecordSource = "Select * From CerrarPedidoMateriaPrimaClientes Order By FechaRecepcion"
                    DataCerrarPedidos.Refresh
                    DGridCerrarPedidos.Refresh
                    TabDepartamentos.Tab = 1
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameCerrarPedidos.Enabled = True
         DataCerrarPedidos.Visible = False
         DGridCerrarPedidos.Visible = False
         FrameBusquedadeDatos.Visible = False
         MskSaldoPedido.Visible = True
         LblSaldo.Visible = True
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameCerrarPedidos.Enabled = False
         DataCerrarPedidos.Visible = True
         DGridCerrarPedidos.Visible = True
         FrameBusquedadeDatos.Visible = True
         MskSaldoPedido.Visible = False
         LblSaldo.Visible = False
    End If
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BMateriaPrima = True Then
            Txttexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
            'MskCan.Text = DBGridBusqueda.Columns(1).Text
            Txttexto.Item(1).SetFocus
        ElseIf BPedido = True Then
            Txttexto.Item(2).Text = DBGridBusqueda.Columns(1).Text
            Txttexto.Item(2).SetFocus
        End If
            FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                Txttexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                'MskCan.Text = DBGridBusqueda.Columns(1).Text
                Txttexto.Item(1).SetFocus
            ElseIf BPedido = True Then
                Txttexto.Item(2).Text = DBGridBusqueda.Columns(1).Text
                Txttexto.Item(2).SetFocus
            End If
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub DGridCerrarPedidos_HeadClick(ByVal ColIndex As Integer)
        DataCerrarPedidos.RecordSource = "Select * from CerrarPedidoMateriaPrimaClientes Order by " & DGridCerrarPedidos.Columns(ColIndex).DataField
        DataCerrarPedidos.Refresh
        DGridCerrarPedidos.Refresh
End Sub

Private Sub Form_Load()
    DataCerrarPedidos.Connect = GConnect
    DataBusqueda.Connect = GConnect
    
    DataCerrarPedidos.DatabaseName = BasedeDatos
    DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub MskCan_GotFocus()
        MskCan.SelStart = 0
        MskCan.SelLength = Len(MskCan.Text)
End Sub

Private Sub MskCan_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskFecRec_GotFocus()
            MskFecRec.SelStart = 0
            MskFecRec.SelLength = Len(MskFecRec.Text)
End Sub

Private Sub MskFecRec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            lblfecini.Visible = False
            LblFecFin.Visible = False
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusqueda.Caption = "No. Despacho"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 1 Then
            lblfecini.Visible = True
            LblFecFin.Visible = True
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusqueda.Caption = ""
            TxtBusqueda.Visible = False
            DtpFecIni.SetFocus
    ElseIf Index = 2 Then
            lblfecini.Visible = False
            LblFecFin.Visible = False
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusqueda.Caption = "No. Pedido"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    End If
    
End Sub

Private Sub TabDepartamentos_Click(PreviousTab As Integer)
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date
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
    If Index = 1 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & Txttexto.Item(1).Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblMateriaPrima.Caption = ""
                End If
    ElseIf Index = 2 Then
            Set RBuscaSaldo = Db.OpenRecordset("Select SaldoPorEntregar From PedidosMateriaPrimaClientes Where Documento = '" & Txttexto.Item(2).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                If RBuscaSaldo.RecordCount > 0 Then
                    MskSaldoPedido.Text = RBuscaSaldo!SaldoPorEntregar
                Else
                    MskSaldoPedido.Text = 0
                End If
    End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'BUSCA Y AGRUPA TODOS LOS CODIGOS DE MATERIA PRIMA QUE VINO EN LA RECEPCION DE BODEGA
    If Index = 1 Then
        DataBusqueda.RecordSource = "Select Codigo, Sum(Cantidad) From DetalleEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "' Group By Codigo"
        BMateriaPrima = True
        BPedido = False
    'BUSCA TODOS LOS PEDIDOS DE LA MATERIA PRIMA
    ElseIf Index = 2 Then
        DataBusqueda.RecordSource = "Select P.FechaPedido, P.Documento, P.CantidadPedido, P.CantidadEntregada, P.SaldoPorEntregar, C.Descripcion From PedidosMateriaPrimaClientes As P, Clientes as C Where P.Codigo = '" & Txttexto.Item(1).Text & "' And P.SaldoPorEntregar > 0 And P.Cliente = C.CodigoCliente"
        BMateriaPrima = False
        BPedido = True
    End If
        DataBusqueda.Refresh
        DBGridBusqueda.Refresh
        
                    If Index = 1 Then
                        DBGridBusqueda.Columns(0).Width = 1500
                        DBGridBusqueda.Columns(1).Width = 1500
                        DBGridBusqueda.Columns(0).Caption = "Codigo"
                        DBGridBusqueda.Columns(1).Caption = "Total"
                        DBGridBusqueda.Columns(1).NumberFormat = "#,###,##0"
                    ElseIf Index = 2 Then
                        DBGridBusqueda.Columns(0).Width = 1000
                        DBGridBusqueda.Columns(1).Width = 1200
                        DBGridBusqueda.Columns(2).Width = 1200
                        DBGridBusqueda.Columns(3).Width = 1200
                        DBGridBusqueda.Columns(4).Width = 1200
                        DBGridBusqueda.Columns(5).Width = 1200
                        DBGridBusqueda.Columns(0).Caption = "Fecha"
                        DBGridBusqueda.Columns(1).Caption = "Pedido"
                        DBGridBusqueda.Columns(2).Caption = "Inicio"
                        DBGridBusqueda.Columns(3).Caption = "Entregado"
                        DBGridBusqueda.Columns(4).Caption = "Saldo"
                        DBGridBusqueda.Columns(5).Caption = "Cliente"
                        DBGridBusqueda.Columns(2).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(3).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(4).NumberFormat = "#,###,##0"
                    End If
        FrameBusqueda.Visible = True
        DBGridBusqueda.SetFocus
        
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
    Txttexto.Item(Index).SelStart = 0
    Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
                SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                'BUSCA Y AGRUPA TODOS LOS CODIGOS DE MATERIA PRIMA QUE VINO EN EL DESPACHO DE BODEGA
                If Index = 1 Then
                    DataBusqueda.RecordSource = "Select Codigo, Sum(Cantidad) From DetalleEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "' Group By Codigo"
                    BMateriaPrima = True
                    BPedido = False
                'BUSCA TODOS LOS PEDIDOS DE LA MATERIA PRIMA
                ElseIf Index = 2 Then
                    DataBusqueda.RecordSource = "Select P.FechaPedido, P.Documento, P.CantidadPedido, P.CantidadEntregada, P.SaldoPorEntregar, C.Descripcion From PedidosMateriaPrimaClientes As P, Clientes as C Where P.Codigo = '" & Txttexto.Item(1).Text & "' And P.SaldoPorEntregar > 0 And P.Cliente = C.CodigoCliente"
                    BMateriaPrima = False
                    BPedido = True
                End If
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    
                    If Index = 1 Then
                        DBGridBusqueda.Columns(0).Width = 1500
                        DBGridBusqueda.Columns(1).Width = 1500
                        DBGridBusqueda.Columns(0).Caption = "Codigo"
                        DBGridBusqueda.Columns(1).Caption = "Total"
                        DBGridBusqueda.Columns(1).NumberFormat = "#,###,##0"
                    ElseIf Index = 2 Then
                        DBGridBusqueda.Columns(0).Width = 1000
                        DBGridBusqueda.Columns(1).Width = 1200
                        DBGridBusqueda.Columns(2).Width = 1200
                        DBGridBusqueda.Columns(3).Width = 1200
                        DBGridBusqueda.Columns(4).Width = 1200
                        DBGridBusqueda.Columns(5).Width = 1200
                        DBGridBusqueda.Columns(0).Caption = "Fecha"
                        DBGridBusqueda.Columns(1).Caption = "Pedido"
                        DBGridBusqueda.Columns(2).Caption = "Inicio"
                        DBGridBusqueda.Columns(3).Caption = "Entregado"
                        DBGridBusqueda.Columns(4).Caption = "Saldo"
                        DBGridBusqueda.Columns(5).Caption = "Cliente"
                        DBGridBusqueda.Columns(2).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(3).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(4).NumberFormat = "#,###,##0"
                    End If
                        FrameBusqueda.Visible = True
                        DBGridBusqueda.SetFocus
        End If
        
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
        'NUMERO DE RECEPCION
        If Index = 0 Then
                Set RBuscaDespacho = Db.OpenRecordset("Select Fecha From EncabezadoEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "'")
                    If RBuscaDespacho.RecordCount > 0 Then
                            MskFecRec.Text = RBuscaDespacho!Fecha
                    End If
        End If
        
        'CODIGO MATERIA PRIMA
        If Index = 1 Then
                'BUSCA LA CANTIDAD QUE TRAE EL DESPACHO
                Set RBuscaCantidadRecepcion = Db.OpenRecordset("Select Sum(Cantidad) From DetalleEgresosMateriaPrima Where Documento = '" & Txttexto.Item(0).Text & "' And Codigo = '" & Txttexto.Item(1).Text & "'")
                    If RBuscaCantidadRecepcion.RecordCount > 0 Then
                        If IsNull(RBuscaCantidadRecepcion(0)) Then
                            VCantidadSalida = 0
                        Else
                            VCantidadSalida = RBuscaCantidadRecepcion(0)
                        End If
                    Else
                        VCantidadSalida = 0
                    End If
                
                'BUSCA LA CANTIDAD QUE TIENE EL CIERRE DE PEDIDOS
                Set RBuscaCantidadPedido = Db.OpenRecordset("Select Sum(Cantidad) From CerrarPedidoMateriaPrimaClientes Where Documento = '" & Txttexto.Item(0).Text & "' And CodigoMateriaPrima = '" & Txttexto.Item(1).Text & "'")
                    If RBuscaCantidadPedido.RecordCount > 0 Then
                        If IsNull(RBuscaCantidadPedido(0)) Then
                            VCantidadCerrarPedido = 0
                        Else
                            VCantidadCerrarPedido = RBuscaCantidadPedido(0)
                        End If
                    Else
                        VCantidadCerrarPedido = 0
                    End If
                'ASIGNA LA CANTIDAD QUE QUEDA DE LA RECEPCION PARA APLICARSELA AL PEDIDO
                MskCan.Text = Val(VCantidadSalida) - Val(VCantidadCerrarPedido)
        End If
        
End Sub
