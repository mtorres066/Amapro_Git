VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraBatch 
   BackColor       =   &H000000FF&
   Caption         =   "Genera Batch e Imprime Certificado"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "GeneraBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8265
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
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7080
         Picture         =   "GeneraBatch.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   6735
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   3495
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Genera Batch"
      TabPicture(0)   =   "GeneraBatch.frx":293C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdSalida"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdGenerar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Imprime Certificado"
      TabPicture(1)   =   "GeneraBatch.frx":3216
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmdImp"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TxtFac"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Imprime Certificado x No Folio"
      TabPicture(2)   =   "GeneraBatch.frx":3530
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdSalImp"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdImpFol"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TxtFol"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.TextBox TxtFol 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73800
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdImpFol 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69840
         Picture         =   "GeneraBatch.frx":368A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalImp 
         Caption         =   "&Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -68520
         Picture         =   "GeneraBatch.frx":37D4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69480
         Picture         =   "GeneraBatch.frx":5846
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3120
         Width           =   2235
      End
      Begin VB.TextBox TxtFac 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   -74040
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton CmdImp 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69480
         Picture         =   "GeneraBatch.frx":78B8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   2235
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Del Envase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   24
         Top             =   1680
         Width           =   5055
         Begin VB.TextBox TxtDesLin3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox TxtBatch3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   720
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtLinea3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   7
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Batch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos De La Tapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1455
         Left            =   -74880
         TabIndex        =   20
         Top             =   3120
         Width           =   5055
         Begin VB.TextBox TxtBatch2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   720
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtLinea2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   9
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtDesLin2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   1440
            TabIndex        =   21
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Batch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5400
         Picture         =   "GeneraBatch.frx":7A02
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalida 
         Caption         =   "&Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6720
         Picture         =   "GeneraBatch.frx":9A74
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Del Envase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   2055
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   5055
         Begin VB.TextBox TxtCliDes 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox TxtCli 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   840
            TabIndex        =   2
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TxtLinea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   840
            MaxLength       =   2
            TabIndex        =   1
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtBatch 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   840
            TabIndex        =   0
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtDesLin 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   1560
            TabIndex        =   17
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Batch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Folio"
         Height          =   195
         Left            =   -74520
         TabIndex        =   35
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   795
      End
   End
End
Attribute VB_Name = "GeneraBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaFechaProduccion As New ADODB.Recordset
Dim RBuscaHoraProduccion As New ADODB.Recordset
Dim RBuscaHoraProduccion2 As New ADODB.Recordset
Dim RRutinas As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaUltimaRutina As New ADODB.Recordset
Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaFichaTecnica2 As New ADODB.Recordset
Dim RBuscaCabezales As New ADODB.Recordset
Dim RBatch As New ADODB.Recordset
Dim RBuscaBatch As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaFolio As New ADODB.Recordset
Dim RBuscaClienteConOrden As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset

Dim Cont As Integer
Dim VFichaTecnica As String
Dim VFichaTecnica2 As String
Dim VFondo As String
Dim VTapa As String
Dim VVariables As String
Dim VVariables2 As String
Dim RMedia As New ADODB.Recordset
Dim RDesviacion As New ADODB.Recordset
Dim RBatchDatos As New ADODB.Recordset
Dim RMenor As New ADODB.Recordset
Dim RMayor As New ADODB.Recordset
Dim RVariables As New ADODB.Recordset
Dim RCuentaPaletUnidades As New ADODB.Recordset
Dim RBuscaFondoTapa As New ADODB.Recordset
Dim RBuscamaximo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim LSP As Double
Dim LIP As Double
Dim VCabezales As Long
Dim VMensaje As String

Dim VCV As Single
Dim VMinCliente As Single
Dim VMaxCliente As Single
Dim VCp As Single
Dim VZ1 As Single
Dim VZ2 As Single
Dim VCPK As Single

Dim VTotalPalet As Long
Dim VTotalUnidades As Long
Dim VTotalPalet2 As Long
Dim VTotalUnidades2 As Long
Dim VTexto As String
Dim VFolio As Long
Dim VBatch2 As Long

Dim BCliente As Boolean
Dim BFichaTecnica As Boolean
Dim VOrden As String
Dim VMensaje2 As String


Private Sub CmdGenerar_Click()
On Error Resume Next
        
        Cont = 0
        VVariables = ""
        
        If TxtBatch.Text = "" Then
                MsgBox "Numero de Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
        End If
        
        If Not IsNumeric(TxtBatch.Text) Then
                MsgBox "Numero de Batch Solo Puede Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
        End If
        
        Set RBuscaCliente = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
                    End If
                    If RBuscaCliente.RecordCount > 0 Then
                        
                    Else
                        MsgBox "Cliente No Existe", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                If TxtCli.Text <> "" Then
                Else
                    MsgBox "Cliente No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If
                
                'Set RBuscaFichaTecnica = New ADODB.Recordset
                '    'BUSCAMOS FICHA TECNICA
                '    If GOrigenDeDatos = "AmaproAccess" Then
                '        Call Abrir_Recordset(RBuscaFichaTecnica, "Select Orden From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                '    Else 'ORACLE
                '        Call Abrir_Recordset(RBuscaFichaTecnica, "Select Orden From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                '    End If
                '
                '    If RBuscaFichaTecnica.RecordCount > 0 Then
                '        VOrden = RBuscaFichaTecnica(0)
                '    Else
                '        MsgBox "No Tiene Asignada La Orden El Batch Y Linea De Produccion", vbOKOnly + vbInformation, "Informacion"
                '        Exit Sub
                '    End If
                '
                'Set RBuscaClienteConOrden = New ADODB.Recordset
                '    If GOrigenDeDatos = "AmaproAccess" Then
                '        Call Abrir_Recordset(RBuscaClienteConOrden, "Select Cliente From EncabezadoOrdenProduccion Where Documento = '" & VOrden & "'")
                '    Else 'ORACLE
                '        Call Abrir_Recordset(RBuscaClienteConOrden, "Select Cliente From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(VOrden) & "'")
                '    End If
                '    If RBuscaClienteConOrden.RecordCount > 0 Then
                '        If TxtCli.Text = RBuscaClienteConOrden!Cliente Then
                '        Else
                '            VMensaje = MsgBox("Cliente No Corresponde Al Cliente De La Orden De Produccion, Desea Grabar?", vbYesNo, "Informacion")
                '            If VMensaje = vbYes Then
                '                VMensaje2 = InputBox("Ingrese Clave De Autorizacion", "Informacion")
                '                If VMensaje2 = "1" Then
                '                Else
                '                    MsgBox "Clave Incorrecta", vbOKOnly + vbInformation, "Informacion"
                '                    Exit Sub
                '                End If
                '            Else
                '                Exit Sub
                '            End If
                '        End If
                '    Else
                '        MsgBox "Numero De Orden De Produccion No Esta Abierta En Las Ordenes De Produccion", vbOKOnly + vbInformation, "Informacion"
                '        Exit Sub
                '    End If
        

        
        'VERIFICA SI YA EXISTE EL BATCH
        Set RBuscaBatch = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBatch, "Select * From BatchDatos Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBatch, "Select * From BatchDatos Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
            End If
        
        If RBuscaBatch.RecordCount > 0 Then
                VMensaje = MsgBox("Batch Ya Se Genero, Desea Generarlo Otra Vez", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If VMensaje = vbYes Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Conexion.Execute ("Delete from batch Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                        Conexion.Execute ("Delete from batchdatos Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                    Else
                        Conexion.Execute ("Delete from batch Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                        Conexion.Execute ("Delete from batchdatos Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                    End If
                Else
                    Exit Sub
                End If
        End If
        
        MousePointer = 11
        
        
        Set RBuscaFichaTecnica = New ADODB.Recordset
        'BUSCAMOS FICHA TECNICA
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
        End If
        
        If RBuscaFichaTecnica.RecordCount > 0 Then
            VFichaTecnica = RBuscaFichaTecnica(0)
        Else
            Set RBuscaFichaTecnica2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Esp_Tec, Linea From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Esp_Tec, Linea From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
                        If RBuscaFichaTecnica2.RecordCount > 0 Then
                            VFichaTecnica = RBuscaFichaTecnica2(0)
                        Else
                            MsgBox "Batch No Existe Verifique en Produccion Interna o Liberada", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            Exit Sub
                        End If
        End If
        
        
        
        'BUSCA QUE CATALOGO DE VARIABLES TIENE LA FICHA TECNICA
        Set RVariables = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RVariables, "Select Variables From FichaTecnica Where Esp_Tec = '" & VFichaTecnica & "'")
        Else
            Call Abrir_Recordset(RVariables, "Select Variables From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica) & "'")
        End If
            If RVariables.RecordCount > 0 Then
                VVariables = RVariables(0)
            Else
                VVariables = ""
            End If
        
        'Agrega Datos Tabla BATCH PARA LUEGO SACAR LOS DATOS ESTADISTICOS DE AQUI
        Batch
            
            
    '_________________________________________________________________________________________________________
            
            
            
            
            
            'GRABA DATOS DE DENSIDAD
            'BUSCAMOS LAS RUTINAS QUE SI SE IMPRIMEN Y LOS CABEZALES QUE TIENE CADA RUTINA
            
            Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And V.Codigo = '" & UCase(VVariables) & "'")
            Else
                Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And UPPER(V.Codigo) = '" & UCase(VVariables) & "'")
            End If
                       
            If RRutinas.RecordCount > 0 Then
                         Do Until RRutinas.EOF
                         
                                 'SACA LA MEDIA DEL BATCH
                                 Set RMedia = New ADODB.Recordset
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                         Call Abrir_Recordset(RMedia, "Select Avg(valor) From Batch Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                                     Else 'ORACLE
                                         Call Abrir_Recordset(RMedia, "Select Avg(valor) From Batch Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                     End If
                                 
                                 'SACA LA DESVIACION DEL BATCH
                                 Set RDesviacion = New ADODB.Recordset
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                         Call Abrir_Recordset(RDesviacion, "Select Stdevp(valor) From Batch Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                                     Else 'ORACLE
                                         Call Abrir_Recordset(RDesviacion, "Select Stddev(valor) From Batch Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                     End If
                                 
                                 'SACA LA DESVIACION DEL BATCH
                                 Set RMenor = New ADODB.Recordset
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                         Call Abrir_Recordset(RMenor, "Select Min(valor) From Batch Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' and Rutina = '" & RRutinas!Rutina & "'")
                                     Else 'ORACLE
                                         Call Abrir_Recordset(RMenor, "Select Min(valor) From Batch Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' and UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                     End If
                         
                                 'SACA LA DESVIACION DEL BATCH
                                 Set RMayor = New ADODB.Recordset
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                         Call Abrir_Recordset(RMayor, "Select Max(valor) From Batch Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                                     Else 'ORACLE
                                         Call Abrir_Recordset(RMayor, "Select Max(valor) From Batch Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                     End If
                                 
                                 'SI ENCUENTRA DATOS
                                 If Not IsNull(RMedia(0)) Then
                                                 
                                                 If RDesviacion(0) = 0 Then
                                                     VCV = 0
                                                 Else
                                                     If RMedia(0) > 0 Then
                                                         VCV = ((RDesviacion(0) / RMedia(0)) * 100)
                                                     Else
                                                         VCV = 0
                                                     End If
                                                 
                                                 End If
                                                 
                                                 LSP = RMedia(0) + 3 * RDesviacion(0)
                                                 LIP = RMedia(0) - 3 * RDesviacion(0)
                                                 
                                                 
                                                 Set RVariables = New ADODB.Recordset
                                                     If GOrigenDeDatos = "AmaproAccess" Then
                                                         Call Abrir_Recordset(RVariables, "Select MinimoClienteMilimetros, MaximoClienteMilimetros From VariablesMedia Where Codigo = '" & VVariables & "' and Rutina = '" & RRutinas!Rutina & "'")
                                                     Else 'ORACLE
                                                         Call Abrir_Recordset(RVariables, "Select MinimoClienteMilimetros, MaximoClienteMilimetros From VariablesMedia Where UPPER(Codigo) = '" & UCase(VVariables) & "' and UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                                     End If
                                                     
                                                     If RVariables.RecordCount > 0 Then
                                                         VMinCliente = RVariables(0)
                                                         VMaxCliente = RVariables(1)
                                                     Else
                                                         VMinCliente = "0"
                                                         VMaxCliente = "0"
                                                     End If
                                                                                    
                                                 If RDesviacion(0) = 0 Then
                                                     VCp = 0
                                                     
                                                 Else
                                                     If RVariables.RecordCount > 0 Then
                                                         VCp = ((RVariables(1) - RVariables(0)) / (LSP - LIP))
                                                         VMinCliente = RVariables(0)
                                                         VMaxCliente = RVariables(1)
                                                     Else
                                                         VCp = 0
                                                         VMinCliente = 0
                                                         VMaxCliente = 0
                                                         
                                                     End If
                                                 End If
                                                 
                                                 'SACA LA DIFERENCIA DE LA ESPECIFICACION MAXIMA CONTRA LA MEDIA
                                                 If RDesviacion(0) = 0 Then
                                                     VZ1 = (VMaxCliente - RMedia(0))
                                                 Else
                                                     VZ1 = ((VMaxCliente - RMedia(0)) / RDesviacion(0))
                                                 End If
                                                 'SACA LA DIFERENCIA DE LA MEDIA CONTRA AL ESPECIFICACION MINIMA
                                                 If RDesviacion(0) = 0 Then
                                                     VZ2 = (RMedia(0) - VMinCliente)
                                                 Else
                                                     VZ2 = ((RMedia(0) - VMinCliente) / RDesviacion(0))
                                                 End If
                                                 'EL DATO MENOR SE DIVIDE EN 3
                                                 If VZ1 < VZ2 Then
                                                     VCPK = VZ1 / 3
                                                 Else
                                                     VCPK = VZ2 / 3
                                                 End If
                                                                                     
                                     
                                         
                                         Conexion.Execute "Insert Into BatchDatos (Batch, Rutina, Lim_Pro_In, Lim_Pro_Su, Cv, Lim_Esp_In, Lim_Esp_Su, Cp, Des_Std, Media, Dat_Men, Dat_May, Linea, Cpk, FichaTecnica, BatchUnificado, Cliente) VALUES(" & TxtBatch & ", '" & RRutinas!Rutina & "', " & LIP & ", " & LSP & ", " & VCV & ", " & VMinCliente & ", " & VMaxCliente & ", " & VCp & ", " & RDesviacion(0) & ", " & RMedia(0) & ", " & RMenor(0) & ", " & RMayor(0) & ", '" & TxtLinea.Text & "', " & VCPK & ", '" & VFichaTecnica & "', 0, '" & TxtCli.Text & "')"
                                         If Err <> 0 Then
                                             MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                             Err.Clear
                                         End If
                                 End If
                         
                             RRutinas.MoveNext
                        Loop
                    Else
                        MsgBox "No Hay Rutinas Capturadas", vbOKOnly + vbInformation, "Informacion"
                    End If
            
    MousePointer = 0
    
    
           If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Exit Sub
           End If
           
           
            
           
    MsgBox "Batch Generado Con Exito", vbOKOnly + vbInformation, "Informacion"

End Sub


Private Sub CmdImp_Click()
        
        If TxtBatch3.Text = "" Then
                MsgBox "Numero de Batch Del Envase Incorrecto", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
        End If
        
        If Not IsNumeric(TxtBatch3.Text) Then
                MsgBox "Numero de Batch De Envase Solo Puede Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
        End If
        
        If TxtBatch2.Text <> "" And TxtLinea2.Text <> "" Then
                If TxtBatch2.Text = "" Then
                        MsgBox "Numero de Batch De La Tapa Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                
                If Not IsNumeric(TxtBatch2.Text) Then
                        MsgBox "Numero de Batch De La Tapa Solo Puede Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                VBatch2 = TxtBatch2.Text
        Else
            VBatch2 = 0
        
        End If
        
        
        
        Set RBuscaFichaTecnica = New ADODB.Recordset
        'BUSCAMOS FICHA TECNICA
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch3.Text & " And Linea = '" & TxtLinea3.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch3.Text & " And UPPER(Linea) = '" & UCase(TxtLinea3.Text) & "'")
        End If
        
        If RBuscaFichaTecnica.RecordCount > 0 Then
            VFichaTecnica = RBuscaFichaTecnica(0)
        Else
            VFichaTecnica = ""
        End If
        
        'BUSCA EL FONDO Y LA TAPA
        Set RBuscaFondoTapa = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFondoTapa, "Select Fondo, Tapa From FichaTecnica Where Esp_Tec = '" & VFichaTecnica & "'")
        Else
            Call Abrir_Recordset(RBuscaFondoTapa, "Select Fondo, Tapa From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica) & "'")
        End If
            If RBuscaFondoTapa.RecordCount > 0 Then
                If IsNull(RBuscaFondoTapa!Fondo) Then
                    VFondo = ""
                ElseIf IsNull(RBuscaFondoTapa!Tapa) Then
                    VTapa = ""
                Else
                    VFondo = RBuscaFondoTapa!Fondo
                    VTapa = RBuscaFondoTapa!Tapa
                End If
            Else
                VFondo = ""
                VTapa = ""
            End If
        
        'CUENTA LOS PALETS Y UNIDADES
           Set RCuentaPaletUnidades = New ADODB.Recordset
            Call Abrir_Recordset(RCuentaPaletUnidades, "Select Count(*), Sum(Envases) From Produccion Where Batch = " & TxtBatch3.Text & " And Linea = '" & TxtLinea3.Text & "' And (Calidad = 'A' OR Calidad = 'I')")
                If RCuentaPaletUnidades.RecordCount > 0 Then
                    If IsNull(RCuentaPaletUnidades(0)) Then
                        VTotalPalet = "0"
                        VTotalUnidades = "0"
                    Else
                        VTotalPalet = RCuentaPaletUnidades(0)
                        VTotalUnidades = RCuentaPaletUnidades(1)
                    End If
                Else
                    VTotalPalet = "0"
                    VTotalUnidades = "0"
                End If
                
           If TxtBatch2.Text <> "" And TxtLinea2.Text <> "" Then
                   'CUENTA LOS PALETS Y UNIDADES
                   Set RCuentaPaletUnidades = New ADODB.Recordset
                    Call Abrir_Recordset(RCuentaPaletUnidades, "Select Count(*), Sum(Envases) From Produccion Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And (Calidad = 'A' OR Calidad = 'I')")
                        If RCuentaPaletUnidades.RecordCount > 0 Then
                            If IsNull(RCuentaPaletUnidades(0)) Then
                                VTotalPalet2 = "0"
                                VTotalUnidades2 = "0"
                            Else
                                VTotalPalet2 = RCuentaPaletUnidades(0)
                                VTotalUnidades2 = RCuentaPaletUnidades(1)
                            End If
                        Else
                            VTotalPalet2 = "0"
                            VTotalUnidades2 = "0"
                        End If
            Else
                    VTotalPalet2 = "0"
                    VTotalUnidades2 = "0"
            End If
            
            
            
                
            Set RBuscamaximo = New ADODB.Recordset
                Call Abrir_Recordset(RBuscamaximo, "Select Max(Folio) from BatchInformacion")
                    If RBuscamaximo.RecordCount > 0 Then
                        If IsNull(RBuscamaximo(0)) Then
                            VFolio = "1"
                        Else
                            VFolio = Val(RBuscamaximo(0)) + 1
                        End If
                    Else
                        VFolio = "1"
                    End If
           
            Set RBuscaCliente = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaCliente, "Select Cliente From BatchDatos Where Batch = " & TxtBatch3.Text & " And Linea = '" & TxtLinea3.Text & "'")
                    If RBuscaCliente.RecordCount > 0 Then
                        
                    Else
                        MsgBox "Cliente Que Tiene Asignado El Batch No Existe", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                
           
            VTexto = TxtBatch3 & ", '"
            VTexto = VTexto & TxtLinea3.Text & "', "
            VTexto = VTexto & VTotalPalet & ", "
            VTexto = VTexto & VTotalUnidades & ", '"
            VTexto = VTexto & RBuscaCliente!Cliente & "', '" 'CLIENTE
            VTexto = VTexto & TxtFac.Text & "', '" ' FACTURA
            VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
            VTexto = VTexto & VFondo & "', '" 'FONDO
            VTexto = VTexto & VTapa & "', " 'TAPA
            VTexto = VTexto & VFolio & ", " 'FOLIO
            VTexto = VTexto & VBatch2 & ", '" 'BATCH 2
            VTexto = VTexto & TxtLinea2.Text & "', " 'lINEA 2
            VTexto = VTexto & VTotalPalet2 & ", " 'TOTAL PALET
            VTexto = VTexto & VTotalUnidades2 'UNIDADES
                   
            'BORRA SI YA EXISTE INFORMACION
            Conexion.Execute "Delete * From BatchInformacion Where Batch = " & TxtBatch3.Text & " And Linea = '" & TxtLinea3.Text & "'"
            If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
            End If
            'AGREGA EL TOTAL DE TARIMAS Y UNIDADES
            Conexion.Execute "Insert Into BatchInformacion Values(" & VTexto & ")"
            If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
            End If
           
            If TxtBatch3.Text <> "" And TxtLinea3.Text <> "" Then
                    'ASIGNA EL NUMERO DE BATCH UNIFICADO PARA PODER IMPRIMIR UNO SOLO
                    Conexion.Execute "Update BatchDatos Set BatchUnificado = " & TxtBatch3.Text & ", LineaUnificada = '" & TxtLinea3.Text & "' Where Batch = " & TxtBatch3.Text & " And Linea = '" & TxtLinea3.Text & "'"
            End If
            
            If TxtBatch2.Text <> "" And TxtLinea2.Text <> "" Then
                    'ASIGNA EL NUMERO DE BATCH UNIFICADO PARA PODER IMPRIMIR UNO SOLO
                    Conexion.Execute "Update BatchDatos Set BatchUnificado = " & TxtBatch3.Text & ", LineaUnificada = '" & TxtLinea3.Text & "' Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'"
            End If
                                            
                
                GCriteriaReporte = "{BatchInformacion.Batch} = " & TxtBatch3.Text & " And {BatchInformacion.Linea} = '" & TxtLinea3.Text & "'"
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "Batch.rpt"
                Else
                    GNombreReporte = "BatchO.rpt"
                End If
                    FrmReporte.Show
        
        
        
        
End Sub

Private Sub CmdImpFol_Click()
            If Not IsNumeric(TxtFol.Text) Then
                MsgBox "Numero De Folio Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                TxtFol.SetFocus
                Exit Sub
            End If
            
            Set RBuscaFolio = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaFolio, "Select folio from batchinformacion Where Folio = " & TxtFol.Text)
                    If RBuscaFolio.RecordCount > 0 Then
                    Else
                        MsgBox "Numero De Folio No Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtFol.SetFocus
                        Exit Sub
                    End If
                
                
                GCriteriaReporte = "{BatchInformacion.Folio} = " & TxtFol.Text
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "Batch.rpt"
                Else
                    GNombreReporte = "BatchO.rpt"
                End If
                    FrmReporte.Show
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdSalImp_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
            If BCliente = True Then
                TxtCli.Text = DbGridBusqueda.Columns(0).Text
                TxtCli.SetFocus
            Else
                
            End If
                FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BCliente = True Then
                TxtCli.Text = DbGridBusqueda.Columns(0).Text
                TxtCli.SetFocus
            Else
                
            End If
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
        If SSTab1.Tab = 0 Then
            TxtBatch.SetFocus
        ElseIf SSTab1.Tab = 1 Then
            TxtCli.SetFocus
        ElseIf SSTab1.Tab = 2 Then
            TxtFol.SetFocus
        End If
        
End Sub


Private Sub TxtBatch_GotFocus()
    TxtBatch.SelStart = 0
    TxtBatch.SelLength = Len(TxtBatch.Text)
End Sub

Private Sub TxtBatch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TxtBatch2_GotFocus()
        TxtBatch2.SelStart = 0
        TxtBatch2.SelLength = Len(TxtBatch2.Text)
End Sub

Private Sub TxtBatch2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtBatch3_GotFocus()
        TxtBatch3.SelStart = 0
        TxtBatch3.SelLength = Len(TxtBatch3.Text)
End Sub

Private Sub TxtBatch3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
    'MATERIA PRIMA
    If BCliente = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where CodigoCliente Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes where UPPER(CodigoCliente) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    'FICHA TECNICA
    ElseIf BFichaTecnica = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    End If
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"

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
        'BUSCA LINEA
        Set RBuscaCliente = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
            End If
            If RBuscaCliente.RecordCount > 0 Then
                TxtCliDes.Text = RBuscaCliente(0)
            Else
                TxtCliDes.Text = ""
            End If

End Sub

Private Sub TxtCli_DblClick()
            Set RBusqueda = New ADODB.Recordset
                BFichaTecnica = False
                BCliente = True
                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "3000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus

End Sub

Private Sub TxtCli_GotFocus()
        TxtCli.SelStart = 0
        TxtCli.SelLength = Len(TxtCli.Text)
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
                BFichaTecnica = False
                BCliente = True
                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "3000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
        End If
End Sub

Private Sub TxtFac_GotFocus()
        TxtFac.SelStart = 0
        TxtFac.SelLength = Len(TxtFac.Text)
End Sub

Private Sub TxtFac_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFol_GotFocus()
        TxtFol.SelStart = 0
        TxtFol.SelLength = Len(TxtFol.Text)
End Sub

Private Sub TxtFol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLinea_Change()
        'BUSCA LINEA
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                TxtDesLin.Text = RBuscaLinea(0)
            Else
                TxtDesLin.Text = ""
            End If
End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub


Public Sub Batch()
On Error Resume Next
Set RBuscaFechaProduccion = New ADODB.Recordset
        'BUSCAMOS LA FECHA MAS JOVEN DEL BATCH
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
        End If
            If RBuscaFechaProduccion.RecordCount > 0 Then
                If IsNull(RBuscaFechaProduccion(0)) Then
                    Set RBuscaFechaProduccion = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                    End If
                End If
            Else
                Set RBuscaFechaProduccion = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
            End If
        
        'BUSCAMOS LA HORA MAS JOVEN DEL BATCH Y LA FECHA JOVEN
        Set RBuscaHoraProduccion = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' And Fec_Prd = #" & RBuscaFechaProduccion(0) & "#")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' And Fec_Prd = To_Date('" & RBuscaFechaProduccion(0) & "', 'dd/mm/yyyy')")
        End If
            If RBuscaHoraProduccion.RecordCount > 0 Then
            Else
                Set RBuscaHoraProduccion2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "' And Fec_Prd = #" & RBuscaFechaProduccion(0) & "#")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' And Fec_Prd = To_Date('" & RBuscaFechaProduccion(0) & "', 'dd/mm/yyyy')")
                End If
            End If
        
        'BUSCAMOS LAS RUTINAS QUE SI SE IMPRIMEN Y LOS CABEZALES QUE TIENE CADA RUTINA
        'PERO QUE SEAN DEL CATALOGO DE LA FICHA TECNICA
        Set RRutinas = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And V.Codigo = '" & UCase(VVariables) & "'")
        Else
            Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And UPPER(V.Codigo) = '" & UCase(VVariables) & "'")
        End If
        
                
            If RRutinas.RecordCount > 0 Then
            
                
                        Do Until RRutinas.EOF
                            
                            VCabezales = RRutinas!Cabezal
                                            
                            Set RBuscaUltimaRutina = New ADODB.Recordset
                            'BUSCAMOS LA ULTIMA Y PENULTIMA RUTINA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaUltimaRutina, "Select * From CapturaRutinas Where Esp_Tec = '" & RBuscaFichaTecnica!Esp_Tec & "' AND Rutina = '" & RRutinas!Rutina & "' And Linea = '" & TxtLinea.Text & "' Order By Fec_Rut, Hor_Rut, Cabezal Asc ")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaUltimaRutina, "Select * From CapturaRutinas Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica) & "' AND UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "' And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "' Order By Fec_Rut, Hor_Rut, Cabezal Asc ")
                            End If
                                           
                                        If RBuscaUltimaRutina.RecordCount > 0 Then
                                            'MUEVE AL ULTIMO
                                            RBuscaUltimaRutina.MoveLast
                                            Cont = 0
                                                       
                                            Do Until Cont = (VCabezales * 2)
                                                        'AGREGA LOS DATOS
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            Conexion.Execute "Insert Into Batch Values(" & TxtBatch.Text & ", '" & RBuscaUltimaRutina!Linea & "', #" & Format(RBuscaUltimaRutina!Fec_Rut, "mm/dd/yyyy") & "#, '" & RBuscaUltimaRutina!Hor_rut & "', '" & RBuscaUltimaRutina!Esp_Tec & "', " & RBuscaUltimaRutina!Cabezal & ", '" & RBuscaUltimaRutina!Rutina & "', " & RBuscaUltimaRutina!Valor & ")"
                                                        Else 'ORACLE
                                                            Conexion.Execute "Insert Into Batch Values(" & TxtBatch.Text & ", '" & RBuscaUltimaRutina!Linea & "', To_Date('" & RBuscaUltimaRutina!Fec_Rut & "', 'dd/mm/yyyy')" & ", '" & RBuscaUltimaRutina!Hor_rut & "', '" & RBuscaUltimaRutina!Esp_Tec & "', " & RBuscaUltimaRutina!Cabezal & ", '" & RBuscaUltimaRutina!Rutina & "', " & RBuscaUltimaRutina!Valor & ")"
                                                        End If
                                                        If Err <> 0 Then
                                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                            Err.Clear
                                                        End If
                                                        RBuscaUltimaRutina.MovePrevious
                                                                                             
                                                    Cont = Cont + 1
                                            Loop
                                            
                                        Else
                                            MsgBox "No se Agregaron Los Datos A La Tabla BATCH", vbOKOnly + vbAbortRetryIgnore, "Informacion"
                                        End If
                                            
                            RRutinas.MoveNext
                        Loop
            Else
                MsgBox "No Hay Rutinas Para Sacar Datos Estadisticos", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
            End If
            
        
End Sub

Public Sub GeneraBatchTapa()
On Error Resume Next
            Cont = 0
        VVariables = ""
        
        If TxtBatch2.Text = "" Then
                Exit Sub
        End If
        
        If Not IsNumeric(TxtBatch2.Text) Then
                MsgBox "Numero de Batch De Tapa Solo Puede Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
        End If
        
        'VERIFICA SI YA EXISTE EL BATCH
        Set RBuscaBatch = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBatch, "Select * From BatchDatos Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBatch, "Select * From BatchDatos Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
            End If
        
        If RBuscaBatch.RecordCount > 0 Then
                'VMensaje = MsgBox("Batch Ya Se Genero, Desea Generarlo Otra Vez", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                'If VMensaje = vbYes Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Conexion.Execute ("Delete from batch Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                        Conexion.Execute ("Delete from batchdatos Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                    Else
                        Conexion.Execute ("Delete from batch Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
                        Conexion.Execute ("Delete from batchdatos Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
                    End If
                'Else
                '    Exit Sub
                'End If
        End If
        
        MousePointer = 11
        
        
        Set RBuscaFichaTecnica = New ADODB.Recordset
        'BUSCAMOS FICHA TECNICA
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Esp_Tec, Linea From Produccion Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
        End If
        
        If RBuscaFichaTecnica.RecordCount > 0 Then
            VFichaTecnica2 = RBuscaFichaTecnica(0)
        Else
            Set RBuscaFichaTecnica2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Esp_Tec, Linea From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Esp_Tec, Linea From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
                End If
                        If RBuscaFichaTecnica2.RecordCount > 0 Then
                            VFichaTecnica2 = RBuscaFichaTecnica2(0)
                        Else
                            MsgBox "Batch No Existe Verifique en Produccion Interna o Liberada", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            Exit Sub
                        End If
        End If
        
        'BUSCA EL FONDO Y LA TAPA
        'Set RBuscaFondoTapa = New ADODB.Recordset
        'If GOrigenDeDatos = "AmaproAccess" Then
        '    Call Abrir_Recordset(RBuscaFondoTapa, "Select Fondo, Tapa From FichaTecnica Where Esp_Tec = '" & VFichaTecnica & "'")
        'Else
        '    Call Abrir_Recordset(RBuscaFondoTapa, "Select Fondo, Tapa From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica) & "'")
        'End If
        '    If RBuscaFondoTapa.RecordCount > 0 Then
        '        VFondo = RBuscaFondoTapa!Fondo
        '        VTapa = RBuscaFondoTapa!Tapa
        '    Else
        '        VFondo = ""
        '        VTapa = ""
        '    End If
        
        
        'BUSCA QUE CATALOGO DE VARIABLES TIENE LA FICHA TECNICA
        Set RVariables = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RVariables, "Select Variables From FichaTecnica Where Esp_Tec = '" & VFichaTecnica2 & "'")
        Else
            Call Abrir_Recordset(RVariables, "Select Variables From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica2) & "'")
        End If
            If RVariables.RecordCount > 0 Then
                VVariables2 = RVariables(0)
            Else
                VVariables2 = ""
            End If
        
        'Agrega Datos Tabla BATCH PARA LUEGO SACAR LOS DATOS ESTADISTICOS DE AQUI
        Batch2
           
           
    '_________________________________________________________________________________________________________
       
           
            
            'GRABA DATOS DE DENSIDAD
            'BUSCAMOS LAS RUTINAS QUE SI SE IMPRIMEN Y LOS CABEZALES QUE TIENE CADA RUTINA
            
            Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And V.Codigo = '" & UCase(VVariables2) & "'")
            Else
                Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And UPPER(V.Codigo) = '" & UCase(VVariables2) & "'")
            End If
                       
            Do Until RRutinas.EOF
            
                    'SACA LA MEDIA DEL BATCH
                    Set RMedia = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RMedia, "Select Avg(valor) From Batch Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RMedia, "Select Avg(valor) From Batch Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                        End If
                    
                    'SACA LA DESVIACION DEL BATCH
                    Set RDesviacion = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDesviacion, "Select Stdevp(valor) From Batch Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDesviacion, "Select Stddev(valor) From Batch Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                        End If
                    
                    'SACA LA DESVIACION DEL BATCH
                    Set RMenor = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RMenor, "Select Min(valor) From Batch Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' and Rutina = '" & RRutinas!Rutina & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RMenor, "Select Min(valor) From Batch Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' and UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                        End If
            
                    'SACA LA DESVIACION DEL BATCH
                    Set RMayor = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RMayor, "Select Max(valor) From Batch Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And Rutina = '" & RRutinas!Rutina & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RMayor, "Select Max(valor) From Batch Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' And UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                        End If
                    
                    'SI ENCUENTRA DATOS
                    If Not IsNull(RMedia(0)) Then
                                    
                                    If RDesviacion(0) = 0 Then
                                        VCV = 0
                                    Else
                                        If RMedia(0) > 0 Then
                                            VCV = ((RDesviacion(0) / RMedia(0)) * 100)
                                        Else
                                            VCV = 0
                                        End If
                                    
                                    End If
                                    
                                    LSP = RMedia(0) + 3 * RDesviacion(0)
                                    LIP = RMedia(0) - 3 * RDesviacion(0)
                                    
                                    
                                    Set RVariables = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RVariables, "Select MinimoClienteMilimetros, MaximoClienteMilimetros From VariablesMedia Where Codigo = '" & VVariables2 & "' and Rutina = '" & RRutinas!Rutina & "'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RVariables, "Select MinimoClienteMilimetros, MaximoClienteMilimetros From VariablesMedia Where UPPER(Codigo) = '" & UCase(VVariables2) & "' and UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "'")
                                        End If
                                        
                                        If RVariables.RecordCount > 0 Then
                                            VMinCliente = RVariables(0)
                                            VMaxCliente = RVariables(1)
                                        Else
                                            VMinCliente = "0"
                                            VMaxCliente = "0"
                                        End If
                                                                       
                                    If RDesviacion(0) = 0 Then
                                        VCp = 0
                                        
                                    Else
                                        If RVariables.RecordCount > 0 Then
                                            VCp = ((RVariables(1) - RVariables(0)) / (LSP - LIP))
                                            VMinCliente = RVariables(0)
                                            VMaxCliente = RVariables(1)
                                        Else
                                            VCp = 0
                                            VMinCliente = 0
                                            VMaxCliente = 0
                                            
                                        End If
                                    End If
                                    
                                    'SACA LA DIFERENCIA DE LA ESPECIFICACION MAXIMA CONTRA LA MEDIA
                                    If RDesviacion(0) = 0 Then
                                        VZ1 = (VMaxCliente - RMedia(0))
                                    Else
                                        VZ1 = ((VMaxCliente - RMedia(0)) / RDesviacion(0))
                                    End If
                                    'SACA LA DIFERENCIA DE LA MEDIA CONTRA AL ESPECIFICACION MINIMA
                                    If RDesviacion(0) = 0 Then
                                        VZ2 = (RMedia(0) - VMinCliente)
                                    Else
                                        VZ2 = ((RMedia(0) - VMinCliente) / RDesviacion(0))
                                    End If
                                    'EL DATO MENOR SE DIVIDE EN 3
                                    If VZ1 < VZ2 Then
                                        VCPK = VZ1 / 3
                                    Else
                                        VCPK = VZ2 / 3
                                    End If
                                                                        
                        
                            
                            Conexion.Execute "Insert Into BatchDatos (Batch, Rutina, Lim_Pro_In, Lim_Pro_Su, Cv, Lim_Esp_In, Lim_Esp_Su, Cp, Des_Std, Media, Dat_Men, Dat_May, Linea, Cpk, FichaTecnica, Cliente) VALUES(" & TxtBatch & ", '" & RRutinas!Rutina & "', " & LIP & ", " & LSP & ", " & VCV & ", " & VMinCliente & ", " & VMaxCliente & ", " & VCp & ", " & RDesviacion(0) & ", " & RMedia(0) & ", " & RMenor(0) & ", " & RMayor(0) & ", '" & TxtLinea.Text & "', " & VCPK & ", '" & VFichaTecnica2 & "', '" & TxtCli.Text & "')"
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                Err.Clear
                            End If
                    End If
            
                RRutinas.MoveNext
           Loop
            
    MousePointer = 0
    
    
           If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Exit Sub
           End If
           
           'CUENTA LOS PALETS Y UNIDADES
           Set RCuentaPaletUnidades = New ADODB.Recordset
            Call Abrir_Recordset(RCuentaPaletUnidades, "Select Count(*), Sum(Envases) From Produccion Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                If RCuentaPaletUnidades.RecordCount > 0 Then
                    If IsNull(RCuentaPaletUnidades(0)) Then
                        VTotalPalet2 = "0"
                        VTotalUnidades2 = "0"
                    Else
                        VTotalPalet2 = RCuentaPaletUnidades(0)
                        VTotalUnidades2 = RCuentaPaletUnidades(1)
                    End If
                Else
                    VTotalPalet2 = "0"
                    VTotalUnidades2 = "0"
                End If
                
            'VTexto = TxtBatch & ", '"
            'VTexto = VTexto & TxtLinea.Text & "', "
            'VTexto = VTexto & VTotalPalet & ", "
            'VTexto = VTexto & VTotalUnidades & ", '', '', '" 'CLIENTE 'FACTURA
            'VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
            'VTexto = VTexto & VFondo & "', '" 'FONDO
            'VTexto = VTexto & VTapa & "'" 'TAPA
            
            'BORRA SI YA EXISTE INFORMACION
            'Conexion.Execute "Delete * From BatchInformacion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtLinea.Text & "'"
            'If Err <> 0 Then
            '        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            '        Err.Clear
            'End If
            'AGREGA EL TOTAL DE TARIMAS Y UNIDADES
            'Conexion.Execute "Insert Into BatchInformacion Values(" & VTexto & ")"
            'If Err <> 0 Then
            '        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            '        Err.Clear
            'End If
        
End Sub

Public Sub Batch2()
On Error Resume Next
            Set RBuscaFechaProduccion = New ADODB.Recordset
        'BUSCAMOS LA FECHA MAS JOVEN DEL BATCH
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From Produccion Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From Produccion Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
        End If
            If RBuscaFechaProduccion.RecordCount > 0 Then
                If IsNull(RBuscaFechaProduccion(0)) Then
                    Set RBuscaFechaProduccion = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
                    End If
                End If
            Else
                Set RBuscaFechaProduccion = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFechaProduccion, "Select Min(Fec_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
                End If
            End If
        
        'BUSCAMOS LA HORA MAS JOVEN DEL BATCH Y LA FECHA JOVEN
        Set RBuscaHoraProduccion = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From Produccion Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And Fec_Prd = #" & RBuscaFechaProduccion(0) & "#")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From Produccion Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' And Fec_Prd = To_Date('" & RBuscaFechaProduccion(0) & "', 'dd/mm/yyyy')")
        End If
            If RBuscaHoraProduccion.RecordCount > 0 Then
            Else
                Set RBuscaHoraProduccion2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And Linea = '" & TxtLinea2.Text & "' And Fec_Prd = #" & RBuscaFechaProduccion(0) & "#")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaHoraProduccion, "Select Min(Hor_Prd) From ProduccionLiberada Where Batch = " & TxtBatch2.Text & " And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' And Fec_Prd = To_Date('" & RBuscaFechaProduccion(0) & "', 'dd/mm/yyyy')")
                End If
            End If
        
        'BUSCAMOS LAS RUTINAS QUE SI SE IMPRIMEN Y LOS CABEZALES QUE TIENE CADA RUTINA
        'PERO QUE SEAN DEL CATALOGO DE LA FICHA TECNICA
        Set RRutinas = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And V.Codigo = '" & UCase(VVariables2) & "'")
        Else
            Call Abrir_Recordset(RRutinas, "Select R.Rutina, R.Cabezal From Rutinas R, VariablesMedia V Where R.Imp_Rut = -1 And R.Rutina = V.Rutina And UPPER(V.Codigo) = '" & UCase(VVariables2) & "'")
        End If
        
                
            If RRutinas.RecordCount > 0 Then
            Else
                MousePointer = 0
            End If
                
            Do Until RRutinas.EOF
                
                VCabezales = RRutinas!Cabezal
                                
                Set RBuscaUltimaRutina = New ADODB.Recordset
                'BUSCAMOS LA ULTIMA Y PENULTIMA RUTINA
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaUltimaRutina, "Select * From CapturaRutinas Where Esp_Tec = '" & RBuscaFichaTecnica!Esp_Tec & "' AND Rutina = '" & RRutinas!Rutina & "' And Linea = '" & TxtLinea2.Text & "' Order By Fec_Rut, Hor_Rut, Cabezal Asc ")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaUltimaRutina, "Select * From CapturaRutinas Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnica) & "' AND UPPER(Rutina) = '" & UCase(RRutinas!Rutina) & "' And UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "' Order By Fec_Rut, Hor_Rut, Cabezal Asc ")
                End If
                               
                            If RBuscaUltimaRutina.RecordCount > 0 Then
                                'MUEVE AL ULTIMO
                                RBuscaUltimaRutina.MoveLast
                                Cont = 0
                                           
                                Do Until Cont = (VCabezales * 2)
                                            'AGREGA LOS DATOS
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Conexion.Execute "Insert Into Batch Values(" & TxtBatch2.Text & ", '" & RBuscaUltimaRutina!Linea & "', #" & Format(RBuscaUltimaRutina!Fec_Rut, "mm/dd/yyyy") & "#, '" & RBuscaUltimaRutina!Hor_rut & "', '" & RBuscaUltimaRutina!Esp_Tec & "', " & RBuscaUltimaRutina!Cabezal & ", '" & RBuscaUltimaRutina!Rutina & "', " & RBuscaUltimaRutina!Valor & ")"
                                            Else 'ORACLE
                                                Conexion.Execute "Insert Into Batch Values(" & TxtBatch2.Text & ", '" & RBuscaUltimaRutina!Linea & "', To_Date('" & RBuscaUltimaRutina!Fec_Rut & "', 'dd/mm/yyyy')" & ", '" & RBuscaUltimaRutina!Hor_rut & "', '" & RBuscaUltimaRutina!Esp_Tec & "', " & RBuscaUltimaRutina!Cabezal & ", '" & RBuscaUltimaRutina!Rutina & "', " & RBuscaUltimaRutina!Valor & ")"
                                            End If
                                            If Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                Err.Clear
                                            End If
                                            RBuscaUltimaRutina.MovePrevious
                                                                                 
                                        Cont = Cont + 1
                                Loop
                            Else
                            
                            End If
                                
                RRutinas.MoveNext
            Loop

End Sub

Private Sub TxtLinea2_Change()
        'BUSCA LINEA
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea2.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea2.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                TxtDesLin2.Text = RBuscaLinea(0)
            Else
                TxtDesLin2.Text = ""
            End If

End Sub

Private Sub TxtLinea2_GotFocus()
        TxtLinea2.SelStart = 0
        TxtLinea2.SelLength = Len(TxtLinea2.Text)
End Sub

Private Sub TxtLinea2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtLinea3_Change()
        'BUSCA LINEA
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea3.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea3.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                TxtDesLin3.Text = RBuscaLinea(0)
            Else
                TxtDesLin3.Text = ""
            End If

End Sub

Private Sub TxtLinea3_GotFocus()
        TxtLinea3.SelStart = 0
        TxtLinea3.SelLength = Len(TxtLinea3.Text)
End Sub

Private Sub TxtLinea3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub
