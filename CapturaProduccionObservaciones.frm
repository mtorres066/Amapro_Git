VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CapturaProduccionObservaciones 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Observaciones De La Tarima  De Produccion"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ControlBox      =   0   'False
   Icon            =   "CapturaProduccionObservaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
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
      Height          =   6855
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   24
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "CapturaProduccionObservaciones.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
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
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "CapturaProduccionObservaciones.frx":237C
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "CapturaProduccionObservaciones.frx":2397
         TabIndex        =   22
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin TabDlg.SSTab TabDefectos 
      Height          =   5535
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaProduccionObservaciones.frx":2D71
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameObservaciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionObservaciones.frx":308B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridObservaciones"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionObservaciones.frx":34DD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label2(5)"
      Tab(2).Control(2)=   "Label2(6)"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "TxtBuscar"
      Tab(2).Control(5)=   "CmdBuscar(0)"
      Tab(2).Control(6)=   "CmdBuscar(1)"
      Tab(2).Control(7)=   "DTPFecIni"
      Tab(2).Control(8)=   "DTPFecFin"
      Tab(2).ControlCount=   9
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68160
         TabIndex        =   38
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37522
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -70200
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37522
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "CapturaProduccionObservaciones.frx":392F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "CapturaProduccionObservaciones.frx":3C39
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   14
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2760
         Width           =   2085
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
         TabIndex        =   30
         Top             =   960
         Width           =   2805
         Begin VB.OptionButton OptFechas 
            Caption         =   "Fechas"
            Height          =   225
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptFechasYLinea 
            Caption         =   "Fechas Y Linea"
            Height          =   195
            Left            =   1080
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FrameObservaciones 
         Caption         =   "Observaciones De La Tarima"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   8235
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Fec_Prd"
            DataSource      =   "DataObservaciones"
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Observaciones"
            DataSource      =   "DataObservaciones"
            Height          =   645
            Index           =   4
            Left            =   1200
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Tarima"
            DataSource      =   "DataObservaciones"
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   3
            ToolTipText     =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Linea"
            DataSource      =   "DataObservaciones"
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Esp_Tec"
            DataSource      =   "DataObservaciones"
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Linea"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   975
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
            Left            =   2880
            TabIndex        =   33
            Top             =   1080
            Width           =   5175
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
            Left            =   2880
            TabIndex        =   32
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGridObservaciones 
         Bindings        =   "CapturaProduccionObservaciones.frx":407B
         Height          =   4305
         Left            =   -74880
         OleObjectBlob   =   "CapturaProduccionObservaciones.frx":409B
         TabIndex        =   11
         Top             =   720
         Width           =   8145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   6
         Left            =   -70800
         TabIndex        =   40
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   5
         Left            =   -68760
         TabIndex        =   39
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   31
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.Data DataObservaciones 
      BackColor       =   &H80000014&
      Caption         =   "Captura De Observaciones De Tarima De Produccion"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProduccionConObservaciones"
      Top             =   6480
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "CapturaProduccionObservaciones.frx":4F8E
      Picture         =   "CapturaProduccionObservaciones.frx":53D0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "CapturaProduccionObservaciones.frx":7442
      Picture         =   "CapturaProduccionObservaciones.frx":7884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "CapturaProduccionObservaciones.frx":7DB6
      Picture         =   "CapturaProduccionObservaciones.frx":81F8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "CapturaProduccionObservaciones.frx":872A
      Picture         =   "CapturaProduccionObservaciones.frx":8B6C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "CapturaProduccionObservaciones.frx":909E
      Picture         =   "CapturaProduccionObservaciones.frx":94E0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "CapturaProduccionObservaciones.frx":9A12
      Picture         =   "CapturaProduccionObservaciones.frx":9E54
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1200
   End
End
Attribute VB_Name = "CapturaProduccionObservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VTipo As String
Dim VCantidad As Long

Dim BFichaTecnica As Boolean
Dim BLinea As Boolean

Dim RBuscaFichaTecnica As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaDefectos As Recordset
Dim RBuscaAtributos As Recordset

Dim VSumaCriticos As Long
Dim VSumaMayores As Long
Dim VSumaMenores As Long


Sub botones()
    If Bandera = True Then
         FrameObservaciones.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataObservaciones.Visible = False
         FrameOpciones.Visible = False
         DBGridObservaciones.Visible = False
    Else
         FrameObservaciones.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataObservaciones.Visible = True
         FrameOpciones.Visible = True
         DBGridObservaciones.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataObservaciones.Recordset
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
                    
                    Txttexto.Item(0).Text = VPFicha
                    Txttexto.Item(1).Text = VPFecha
                    Txttexto.Item(2).Text = VPLinea
                    Txttexto.Item(3).Text = VPTarima
                    
                    Txttexto.Item(4).SetFocus
            'EDITAR
            ElseIf Index = 1 Then
                    'EDITA EL REGISTRO
                    .Edit
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    Txttexto.Item(0).SetFocus
            'GRABAR
            ElseIf Index = 2 Then
                   
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "Esta Tarima Ya Tiene Asignada Observaciones", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(0).SetFocus
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
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        DataObservaciones.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataObservaciones.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataObservaciones.Recordset.EOF Then
                        DataObservaciones.Recordset.MoveLast
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
    With DataObservaciones
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptFechas.Value = True Then
                .RecordSource = ("Select * from ProduccionConObservaciones where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd")
                .Refresh
                DBGridObservaciones.Refresh
            ElseIf OptFechasYLinea.Value = True Then
                .RecordSource = ("Select * from ProduccionConObservaciones where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "' Order By Fec_prd")
                .Refresh
                DBGridObservaciones.Refresh
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From ProduccionConObservaciones"
                .Refresh
                DBGridObservaciones.Refresh
        End If
    End With
        TabDefectos.Tab = 1
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
            If BFichaTecnica = True Then
                Txttexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(0).SetFocus
            ElseIf BLinea = True Then
                Txttexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(2).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BFichaTecnica = True Then
                    Txttexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                    Txttexto.Item(0).SetFocus
                ElseIf BLinea = True Then
                    Txttexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                    Txttexto.Item(2).SetFocus
                End If
                    FrameBusqueda.Visible = False
            End If
                    
End Sub

Private Sub dbgridobservaciones_HeadClick(ByVal ColIndex As Integer)
    DataObservaciones.RecordSource = ("Select * from ProduccionConObservaciones order by " & DBGridObservaciones.Columns(ColIndex).DataField)
    DataObservaciones.Refresh
    DBGridObservaciones.Refresh
End Sub


Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        DataObservaciones.Connect = GConnect
        DataBusqueda.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        DataObservaciones.DatabaseName = BasedeDatos
        DataBusqueda.DatabaseName = BasedeDatos
End Sub


Private Sub OptFechas_Click()
        Lbletiqueta.Caption = ""
End Sub

Private Sub OptFechasYlinea_Click()
        Lbletiqueta.Caption = "Linea"
End Sub

Private Sub TabDefectos_Click(PreviousTab As Integer)
        If TabDefectos.Tab = 2 Then
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
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
            
    'LINEA
    If BLinea = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas where Descrip Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas where Descrip Like '*" & TxtBusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas where Linea Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas where Linea Like '*" & TxtBusqueda.Text & "*'"
                End If
            End If
    'FICHA TECNICA
    ElseIf BFichaTecnica = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '" & TxtBusqueda.Text & "*' Order By Esp_tec"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '*" & TxtBusqueda.Text & "*' Order By Esp_Tec"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '" & TxtBusqueda.Text & "*' Order By Esp_Tec"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '*" & TxtBusqueda.Text & "*' Order By Esp_Tec"
                End If
            End If
    
    End If
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"

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
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
            Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica where Esp_Tec = '" & Txttexto.Item(0).Text & "'")
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFichaTecnica.Caption = ""
                End If
        'BUSCA LA DESCRIPCION DE LINEA
        ElseIf Index = 2 Then
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txttexto.Item(2).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        'SI ELIGE FICHA TECNICA
        If Index = 0 Then
            BFichaTecnica = True
            BLinea = False
            DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
        'LINEAS
        ElseIf Index = 2 Then
            BFichaTecnica = False
            BLinea = True
            DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas"
        End If
        
        If (Index = 0 Or Index = 2) Then
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
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
    
    If KeyAscii = 43 Then
        'SI ELIGE FICHA TECNICA
        If Index = 0 Then
            BFichaTecnica = True
            BLinea = False
            DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
        'LINEAS
        ElseIf Index = 2 Then
            BFichaTecnica = False
            BLinea = True
            DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas"
        End If
        
        If (Index = 0 Or Index = 2) Then
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
        End If
      
    End If
    
End Sub
