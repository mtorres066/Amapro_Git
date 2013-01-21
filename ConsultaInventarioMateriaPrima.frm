VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ConsultaInventarioMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Consulta De Inventario Materia Prima"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "ConsultaInventarioMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
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
      Height          =   6855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "ConsultaInventarioMateriaPrima.frx":5C12
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "ConsultaInventarioMateriaPrima.frx":5C2D
         TabIndex        =   20
         Top             =   1080
         Width           =   8535
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Busqueda"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   8040
         Picture         =   "ConsultaInventarioMateriaPrima.frx":6607
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton OptSaldo 
         BackColor       =   &H00008000&
         Caption         =   "<= Cero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   32
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton OptSaldo 
         BackColor       =   &H00008000&
         Caption         =   "> Cero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox TxtTipo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Width           =   3375
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00008000&
         Caption         =   "Un Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   24
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H00008000&
         Caption         =   "Todos Tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.OptionButton OptFichaBodega 
      BackColor       =   &H00008000&
      Caption         =   "Bodega Y Codigo"
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
      Left            =   2040
      TabIndex        =   21
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton OptBodega 
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OptFichaTecnica 
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox TxtBodega 
      Appearance      =   0  'Flat
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
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   495
      Left            =   11400
      Picture         =   "ConsultaInventarioMateriaPrima.frx":8679
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   0
      Width           =   495
   End
   Begin MSMask.MaskEdBox MskTotalEnvases 
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTotalTarimas 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin VB.Data DataMateriaPrima 
      Caption         =   "Ficha Tecnica"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGridFichaTecnica 
      Bindings        =   "ConsultaInventarioMateriaPrima.frx":A6EB
      Height          =   7335
      Left            =   120
      OleObjectBlob   =   "ConsultaInventarioMateriaPrima.frx":A70A
      TabIndex        =   19
      Top             =   1200
      Width           =   11655
   End
   Begin MSMask.MaskEdBox MskTotalPeso 
      Height          =   285
      Left            =   9000
      TabIndex        =   29
      Top             =   120
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   8520
      TabIndex        =   28
      Top             =   120
      Width           =   435
   End
   Begin VB.Label LblTipo2 
      BackColor       =   &H00008000&
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
      Left            =   6480
      TabIndex        =   27
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label LblTipo 
      BackColor       =   &H00008000&
      Caption         =   "Tipo"
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
      Left            =   4080
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Tarimas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   18
      Top             =   120
      Width           =   675
   End
   Begin VB.Label LblBodega 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
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
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label LblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
   Begin MSForms.CommandButton CmdGenera 
      Default         =   -1  'True
      Height          =   495
      Left            =   10800
      TabIndex        =   1
      ToolTipText     =   "Generar Datos"
      Top             =   0
      Width           =   495
      BackColor       =   12632256
      PicturePosition =   327683
      Size            =   "873;873"
      Picture         =   "ConsultaInventarioMateriaPrima.frx":B100
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "ConsultaInventarioMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaTipo As Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean

Dim BBodega As Boolean
Dim BTipo As Boolean



Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11

            
            If OptFichaTecnica.Value = True Then
                    If OptTodos.Value = True Then
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima As F Where DE.Codigo = F.CodigoMateriaPrima And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad) From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima As F Where DE.Codigo = F.CodigoMateriaPrima Group By DE.Codigo, F.Descripcion Having Sum(De.SaldoDisponibilidad) <= 0"
                        End If
                    Else
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima As F Where DE.Codigo = F.CodigoMateriaPrima And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad) From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima As F Where DE.Codigo = F.CodigoMateriaPrima And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' Group By DE.Codigo, F.Descripcion Having sum(DE.SaldoDisponibilidad) <= 0"
                        End If
                    End If
            ElseIf OptBodega.Value = True Then
                    If OptTodos.Value = True Then
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima Where DE.BodegaDisponibilidad = '" & TxtBodega.Text & "' And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad) From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima Where DE.BodegaDisponibilidad = '" & TxtBodega.Text & "' Group By DE.Codigo, F.Descripcion Having Sum(DE.SaldoDisponibilidad) <= 0"
                        End If
                    Else
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima Where DE.BodegaDisponibilidad = '" & TxtBodega.Text & "' And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad) From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima Where DE.BodegaDisponibilidad = '" & TxtBodega.Text & "' And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' Group By DE.Codigo, F.Descripcion Having Sum(DE.SaldoDisponibilidad) <= 0"
                        End If
                    End If
            ElseIf OptFichaBodega.Value = True Then
                    If OptTodos.Value = True Then
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, De.BodegaDisponibilidad, B.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima, BodegasMateriaPrima as B Where De.BodegaDisponibilidad = B.CodigoBodega And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, De.BodegaDisponibilidad, B.Descripcion, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad), De.BodegaDisponibilidad, B.Descripcion From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima, BodegasMateriaPrima as B Where De.BodegaDisponibilidad = B.CodigoBodega Group By DE.Codigo, F.Descripcion, De.BodegaDisponibilidad, B.Descripcion Having Sum(DE.SaldoDisponibilidad) <= 0"
                        End If
                    Else
                        If OptSaldo.Item(0).Value = True Then
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, De.BodegaDisponibilidad, B.Descripcion, Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), F.UnidadMedida, Sum(De.Peso), F.UnidadMedidaPeso From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima, BodegasMateriaPrima as B Where De.BodegaDisponibilidad = B.CodigoBodega And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0 Group By DE.Codigo, F.Descripcion, F.UnidadMedida, De.BodegaDisponibilidad, B.Descripcion, F.UnidadMedidaPeso"
                        Else
                            DataMateriaPrima.RecordSource = "SELECT DE.Codigo, F.Descripcion, Sum(DE.SaldoDisponibilidad), De.BodegaDisponibilidad, B.Descripcion From DetalleEntradasMateriaPrima As DE INNER JOIN CorrelativosMateriaPrima As F ON DE.Codigo = F.CodigoMateriaPrima, BodegasMateriaPrima as B Where De.BodegaDisponibilidad = B.CodigoBodega And F.TipoDeMateriaPrima = '" & TxtTipo.Text & "' Group By DE.Codigo, F.Descripcion, De.BodegaDisponibilidad, B.Descripcion Having Sum(DE.SaldoDisponibilidad) <= 0"
                        End If
                    End If
            End If
            
            DataMateriaPrima.Refresh
            DBGridFichaTecnica.Refresh
            'POR BODEGA Y FICHA TECNICA
            If OptFichaBodega.Value = True Then
                If OptSaldo.Item(0).Value = True Then
                    DBGridFichaTecnica.Columns(0).Caption = "Codigo"
                    DBGridFichaTecnica.Columns(1).Caption = "Descripcion"
                    DBGridFichaTecnica.Columns(2).Caption = "Bodega"
                    DBGridFichaTecnica.Columns(3).Caption = "Descripcion"
                    DBGridFichaTecnica.Columns(4).Caption = "Bultos"
                    DBGridFichaTecnica.Columns(5).Caption = "Cantidad"
                    DBGridFichaTecnica.Columns(6).Caption = "U/M"
                    DBGridFichaTecnica.Columns(7).Caption = "Peso"
                    DBGridFichaTecnica.Columns(8).Caption = "U/P"
                    
                    DBGridFichaTecnica.Columns(0).Width = "1300"
                    DBGridFichaTecnica.Columns(1).Width = "3500"
                    DBGridFichaTecnica.Columns(2).Width = "500"
                    DBGridFichaTecnica.Columns(3).Width = "1500"
                    DBGridFichaTecnica.Columns(4).Width = "600"
                    DBGridFichaTecnica.Columns(5).NumberFormat = "#,###,##0.00"
                    DBGridFichaTecnica.Columns(5).Width = "1000"
                    DBGridFichaTecnica.Columns(6).Width = "900"
                    DBGridFichaTecnica.Columns(7).NumberFormat = "#,###,##0.00"
                    DBGridFichaTecnica.Columns(7).Width = "1000"
                    DBGridFichaTecnica.Columns(8).Width = "700"
                Else
                    DBGridFichaTecnica.Columns(0).Caption = "Codigo"
                    DBGridFichaTecnica.Columns(1).Caption = "Descripcion"
                    DBGridFichaTecnica.Columns(2).Caption = "Cantidad"
                    DBGridFichaTecnica.Columns(3).Caption = "Bodega"
                    DBGridFichaTecnica.Columns(4).Caption = "Descripcion"
                    
                    DBGridFichaTecnica.Columns(0).Width = "1300"
                    DBGridFichaTecnica.Columns(1).Width = "3500"
                    DBGridFichaTecnica.Columns(2).Width = "1300"
                    DBGridFichaTecnica.Columns(2).NumberFormat = "#,###,##0.00"
                    DBGridFichaTecnica.Columns(3).Width = "500"
                    DBGridFichaTecnica.Columns(4).Width = "3000"
                End If
            Else
                If OptSaldo.Item(0).Value = True Then
                    DBGridFichaTecnica.Columns(0).Caption = "Codigo"
                    DBGridFichaTecnica.Columns(1).Caption = "Descripcion"
                    DBGridFichaTecnica.Columns(2).Caption = "Bultos"
                    DBGridFichaTecnica.Columns(3).Caption = "Cantidad"
                    DBGridFichaTecnica.Columns(3).NumberFormat = "#,###,##0.00"
                    DBGridFichaTecnica.Columns(4).Caption = "U/M"
                    DBGridFichaTecnica.Columns(5).Caption = "Peso"
                    DBGridFichaTecnica.Columns(5).NumberFormat = "#,###,##0.00"
                    DBGridFichaTecnica.Columns(6).Caption = "U/P"
                    DBGridFichaTecnica.Columns(0).Width = "1300"
                    DBGridFichaTecnica.Columns(1).Width = "4000"
                    DBGridFichaTecnica.Columns(2).Width = "1200"
                    DBGridFichaTecnica.Columns(3).Width = "1200"
                    DBGridFichaTecnica.Columns(4).Width = "1000"
                    DBGridFichaTecnica.Columns(5).Width = "1000"
                    DBGridFichaTecnica.Columns(6).Width = "1000"
                Else
                    DBGridFichaTecnica.Columns(0).Caption = "Codigo"
                    DBGridFichaTecnica.Columns(1).Caption = "Descripcion"
                    DBGridFichaTecnica.Columns(2).Caption = "Cantidad"
                    
                    DBGridFichaTecnica.Columns(0).Width = "1300"
                    DBGridFichaTecnica.Columns(1).Width = "3500"
                    DBGridFichaTecnica.Columns(2).Width = "1300"
                    DBGridFichaTecnica.Columns(2).NumberFormat = "#,###,##0.00"
                End If
            End If
            
            
            If OptFichaTecnica.Value = True Then
                If OptTodos.Value = True Then
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad > 0")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad <= 0")
                    End If
                Else
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima as DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima as DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad <= 0")
                    End If
                End If
            ElseIf OptBodega.Value = True Then
                If OptTodos.Value = True Then
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad > 0 And BodegaDisponibilidad = '" & TxtBodega.Text & "'")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad <= 0 And BodegaDisponibilidad = '" & TxtBodega.Text & "'")
                    End If
                Else
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima as DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0 And DE.BodegaDisponibilidad = '" & TxtBodega.Text & "'")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(De.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima as DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad <= 0 And DE.BodegaDisponibilidad = '" & TxtBodega.Text & "'")
                    End If
                End If
            ElseIf OptFichaBodega.Value = True Then
                If OptTodos.Value = True Then
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad > 0 ")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(SaldoDisponibilidad), Sum(SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima Where SaldoDisponibilidad <= 0 ")
                    End If
                Else
                    If OptSaldo.Item(0).Value = True Then
                        Set RTotal = Db.OpenRecordset("SELECT Count(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad > 0 ")
                    Else
                        Set RTotal = Db.OpenRecordset("SELECT Count(DE.SaldoDisponibilidad), Sum(DE.SaldoDisponibilidad), Sum(Peso) From DetalleEntradasMateriaPrima As DE, CorrelativosMateriaPrima as C Where DE.Codigo = C.CodigoMateriaPrima And C.TipoDeMateriaPrima = '" & TxtTipo.Text & "' And DE.SaldoDisponibilidad <= 0 ")
                    End If
                End If
            End If
                        
            If RTotal.RecordCount > 0 Then
                If OptSaldo.Item(1).Value = True Then
                    MskTotalTarimas.Text = 0
                    MskTotalEnvases = 0
                    MskTotalPeso = 0
                Else
                        If Not IsNull(RTotal(0)) Then
                            MskTotalTarimas.Text = RTotal(0)
                            MskTotalEnvases.Text = RTotal(1)
                            MskTotalPeso.Text = RTotal(2)
                        Else
                            MskTotalTarimas.Text = "0"
                            MskTotalEnvases.Text = "0"
                            MskTotalPeso.Text = "0"
                        End If
                End If
            Else
                MskTotalTarimas.Text = 0
                MskTotalEnvases.Text = 0
                MskTotalPeso.Text = 0
            End If
            
            
            If Err <> 0 Then
                'MsgBox Err.Description
            End If
            
MousePointer = 0
        
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False

End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BBodega = True Then
            TxtBodega.Text = DBGridBusqueda.Columns(0)
            TxtBodega.SetFocus
        ElseIf BTipo = True Then
            TxtTipo.Text = DBGridBusqueda.Columns(0)
            TxtTipo.SetFocus
        End If
            FrameBusqueda.Visible = False
            
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                    If BBodega = True Then
                        TxtBodega.Text = DBGridBusqueda.Columns(0)
                        TxtBodega.SetFocus
                    ElseIf BTipo = True Then
                        TxtTipo.Text = DBGridBusqueda.Columns(0)
                        TxtTipo.SetFocus
                    End If
                        FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
            DataMateriaPrima.ConnectionString = GTipoProveedor
            DataBusqueda.ConnectionString = GTipoProveedor

            DataMateriaPrima.Refresh
            DataBusqueda.Refresh
                        
            CmdGenera_Click
End Sub


Private Sub Optbodega_Click()
            LblDescripcion.Visible = True
            TxtBodega.Visible = True
            TxtBodega.SetFocus
End Sub


Private Sub OptFichaBodega_Click()
            LblDescripcion.Visible = False
            LblBodega.Caption = ""
            TxtBodega.Visible = False
End Sub

Private Sub OptFichaTecnica_Click()
            LblDescripcion.Visible = False
            LblBodega.Caption = ""
            TxtBodega.Visible = False
End Sub

Private Sub OptTipo_Click(Index As Integer)
        LblTipo.Visible = True
        TxtTipo.Visible = True
End Sub

Private Sub OptTodos_Click()
        LblTipo.Visible = False
        TxtTipo.Text = ""
        TxtTipo.Visible = False
End Sub

Private Sub Txtbusqueda_Change()
    If BBodega = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoBodega, Descripcion From BodegasMateriaPrima where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoBodega, Descripcion From BodegasMateriaPrima where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoBodega, Descripcion From BodegasMateriaPrima where CodigoBodega Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoBodega, Descripcion From BodegasMateriaPrima where CodigoBodega Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
    Else
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoTipo, Descripcion From TiposDeMateriaPrima where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoTipo, Descripcion From TiposDeMateriaPrima where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoTipo, Descripcion From TiposDeMateriaPrima where CodigoTipo Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select codigoTipo, Descripcion From TiposDeMateriaPrima where CodigoTipo Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
    End If
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBodega_Change()
        If OptBodega.Value = True Then
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodega.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBodega.Caption = RBuscaBodega!Descripcion
                Else
                    LblBodega.Caption = ""
                End If
        End If
            
End Sub

Private Sub TxtBodega_DblClick()
                    BBodega = True
                    BTipo = False
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
End Sub

Private Sub TxtBodega_GotFocus()
        TxtBodega.SelStart = 0
        TxtBodega.SelLength = Len(TxtBodega.Text)
End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    BBodega = True
                    BTipo = False
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
End Sub

Private Sub TxtTipo_Change()
        Set RBuscaTipo = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima WHere CodigoTipo = '" & TxtTipo.Text & "'")
            If RBuscaTipo.RecordCount > 0 Then
                LblTipo2.Caption = RBuscaTipo!Descripcion
            Else
                LblTipo2.Caption = ""
            End If
        
End Sub

Private Sub TxtTipo_DblClick()
                    BBodega = False
                    BTipo = True
                    DataBusqueda.RecordSource = "Select CodigoTipo, Descripcion From TiposDeMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
End Sub

Private Sub TxtTipo_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                    
                If KeyAscii = 43 Then
                    BBodega = False
                    BTipo = True
                    DataBusqueda.RecordSource = "Select CodigoTipo, Descripcion From TiposDeMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
                End If
End Sub
