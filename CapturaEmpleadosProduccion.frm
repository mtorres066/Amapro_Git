VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CapturaEmpleadosProduccion 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Horas Extras"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "CapturaEmpleadosProduccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   5535
      Left            =   6960
      TabIndex        =   29
      Top             =   7920
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Data DataConsultas 
         Caption         =   "Defectos"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   3960
         TabIndex        =   34
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   7560
         Picture         =   "CapturaEmpleadosProduccion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CapturaEmpleadosProduccion.frx":0D0C
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "CapturaEmpleadosProduccion.frx":0D28
         TabIndex        =   37
         Top             =   1080
         Width           =   8175
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
   End
   Begin TabDlg.SSTab TabEmpleados 
      Height          =   6735
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaEmpleadosProduccion.frx":1703
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEmpleados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaEmpleadosProduccion.frx":1A1D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridEmpleados"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaEmpleadosProduccion.frx":1E6F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdBuscar(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "CmdBuscar(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "DtpFecIni"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DtpFecFin"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -69960
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54919169
         CurrentDate     =   37588
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -72120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54919171
         CurrentDate     =   37588
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "CapturaEmpleadosProduccion.frx":22C1
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "CapturaEmpleadosProduccion.frx":25CB
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Frame FrameEmpleados 
         Caption         =   "Datos Del Empleado"
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
         Height          =   3375
         Left            =   1680
         TabIndex        =   18
         Top             =   1680
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasDobles"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   11
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   45
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDobles"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   44
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   9
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   43
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   8
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   42
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   7
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   41
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   6
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   40
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasLaboradas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   5
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   39
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Usuario"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   2
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasLaboradas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Empleado"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Linea"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   2
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Turno"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   1
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Laboradas"
            Height          =   195
            Index           =   12
            Left            =   3480
            TabIndex        =   52
            Top             =   1800
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Diurnas"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   51
            Top             =   2160
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Diurnas"
            Height          =   195
            Index           =   10
            Left            =   3480
            TabIndex        =   50
            Top             =   2160
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Nocturnas"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   49
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Nocturnas"
            Height          =   195
            Index           =   8
            Left            =   3480
            TabIndex        =   48
            Top             =   2520
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Dobles"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   2880
            Width           =   1440
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Dobles"
            Height          =   195
            Index           =   6
            Left            =   3480
            TabIndex        =   46
            Top             =   2880
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   5
            Left            =   6000
            TabIndex        =   28
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LblPuesto 
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
            Left            =   2520
            TabIndex        =   25
            Top             =   1440
            Width           =   5535
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
            Left            =   2520
            TabIndex        =   24
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Laboradas"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empleado"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   420
         End
      End
      Begin MSDBGrid.DBGrid DBGridEmpleados 
         Bindings        =   "CapturaEmpleadosProduccion.frx":2A0D
         Height          =   5865
         Left            =   -74880
         OleObjectBlob   =   "CapturaEmpleadosProduccion.frx":2A29
         TabIndex        =   12
         Top             =   720
         Width           =   11505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   -70560
         TabIndex        =   27
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   -72720
         TabIndex        =   26
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Data DataEmpleados 
      BackColor       =   &H80000014&
      Caption         =   "Empleados De Produccion"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EmpleadosCapturaHoras"
      Top             =   7800
      Width           =   11595
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   9840
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":470C
      Picture         =   "CapturaEmpleadosProduccion.frx":4B4E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   7920
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":6BC0
      Picture         =   "CapturaEmpleadosProduccion.frx":7002
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   6000
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":7534
      Picture         =   "CapturaEmpleadosProduccion.frx":7976
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   4080
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":7EA8
      Picture         =   "CapturaEmpleadosProduccion.frx":82EA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2160
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":881C
      Picture         =   "CapturaEmpleadosProduccion.frx":8C5E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "CapturaEmpleadosProduccion.frx":9190
      Picture         =   "CapturaEmpleadosProduccion.frx":95D2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1800
   End
End
Attribute VB_Name = "CapturaEmpleadosProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BTurno As Boolean
Dim BLinea As Boolean
Dim BPuesto As Boolean

Dim RBuscaLineas As Recordset
Dim RBuscaPuestos As Recordset



Sub botones()
    If Bandera = True Then
         FrameEmpleados.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         DataEmpleados.Visible = False
         DBGridEmpleados.Visible = False
    Else
         FrameEmpleados.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataEmpleados.Visible = True
         DBGridEmpleados.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataEmpleados.Recordset
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
                    MskFec.Text = Date
                    Txttexto.Item(0).SetFocus
                    Txttexto.Item(2).Text = GUsuario
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
                    Txttexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                   
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "En Esta Fecha y Turno Y Linea y Puesto Ya Existe", vbOKOnly + vbInformation, "Informacion"
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
                        DataEmpleados.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataEmpleados.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataEmpleados.Recordset.EOF Then
                        DataEmpleados.Recordset.MoveLast
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
    With DataEmpleados
        'SELECCIONAR DATOS
        If Index = 0 Then
                .RecordSource = ("Select * from EmpleadosProduccion where Fecha >= #" & DTPFecIni.Value & "# And Fecha <= #" & DTPFecFin.Value & "#")
                .Refresh
                DBGridEmpleados.Refresh
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From EmpleadosProduccion"
                .Refresh
                DBGridEmpleados.Refresh
        End If
    End With
        TabEmpleados.Tab = 1
End Sub


Private Sub DBGridConsultas_DblClick()
        If BTurno = True Then
            Txttexto.Item(0).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(0).SetFocus
        ElseIf BLinea = True Then
            Txttexto.Item(1).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(1).SetFocus
        ElseIf BPuesto = True Then
            Txttexto.Item(3).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(3).SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If BTurno = True Then
            Txttexto.Item(0).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(0).SetFocus
        ElseIf BLinea = True Then
            Txttexto.Item(1).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(1).SetFocus
        ElseIf BPuesto = True Then
            Txttexto.Item(3).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(3).SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBgridempleados_HeadClick(ByVal ColIndex As Integer)
    DataEmpleados.RecordSource = ("Select * from EmpleadosProduccion order by " & DBGridEmpleados.Columns(ColIndex).DataField)
    DataEmpleados.Refresh
    DBGridEmpleados.Refresh
End Sub


Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        DataEmpleados.Connect = GConnect
        DataConsultas.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        DataEmpleados.DatabaseName = BasedeDatos
        DataConsultas.DatabaseName = BasedeDatos

End Sub


Private Sub tabempleados_Click(PreviousTab As Integer)
        If TabEmpleados.Tab = 2 Then
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
        End If
End Sub

Private Sub TxtConsultas_Change()
    'PUESTOS
    If BPuesto = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Puestos Where Descripcion Like '" & TxtConsultas.Text & "*' Order By Descripcion"
            Else
                DataConsultas.RecordSource = "Select * From Puestos Where Descripcion Like '*" & TxtConsultas.Text & "*' Order By Descripcion"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Puestos Where CodigoPuesto Like '" & TxtConsultas.Text & "*' Order By Descripcion"
            Else
                DataConsultas.RecordSource = "Select * From Puestos Where CodigoPuesto Like '*" & TxtConsultas.Text & "*' Order By Descripcion"
            End If
        End If
    'LINEA
    ElseIf BLinea = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Lineas Where Descrip Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select * From Lineas Where Descrip Like '*" & TxtConsultas.Text & "*'"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Lineas Where Linea Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select * From Lineas Where Linea Like '*" & TxtConsultas.Text & "*'"
            End If
        End If
    End If
    DataConsultas.Refresh
    DBGridConsultas.Refresh

End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 1 Then
            Set RBuscaLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txttexto.Item(1).Text & "'")
                If RBuscaLineas.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLineas!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        ElseIf Index = 3 Then
            Set RBuscaPuestos = Db.OpenRecordset("Select Descripcion From Puestos Where CodigoPuesto = '" & Txttexto.Item(3).Text & "'")
                If RBuscaPuestos.RecordCount > 0 Then
                    LblPuesto.Caption = RBuscaPuestos!Descripcion
                Else
                    LblPuesto.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 0 Then
                BTurno = True
                BLinea = False
                BPuesto = False
                DataConsultas.RecordSource = "Select * From Turnos"
        ElseIf Index = 1 Then
                BTurno = False
                BLinea = True
                BPuesto = False
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
        ElseIf Index = 3 Then
                BTurno = False
                BLinea = False
                BPuesto = True
                DataConsultas.RecordSource = "Select CodigoPuesto, Descripcion From Puestos"
        End If
        
        If Index = 0 Or Index = 1 Or Index = 3 Then
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                FrameConsultas.Visible = True
                TxtConsultas.SetFocus
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
                If Index = 0 Then
                        BTurno = True
                        BLinea = False
                        BPuesto = False
                        DataConsultas.RecordSource = "Select * From Turnos"
                ElseIf Index = 1 Then
                        BTurno = False
                        BLinea = True
                        BPuesto = False
                        DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
                ElseIf Index = 3 Then
                        BTurno = False
                        BLinea = False
                        BPuesto = True
                        DataConsultas.RecordSource = "Select CodigoPuesto, Descripcion From Puestos"
                End If
                
                If Index = 0 Or Index = 1 Or Index = 3 Then
                        DataConsultas.Refresh
                        DBGridConsultas.Refresh
                        FrameConsultas.Visible = True
                        TxtConsultas.SetFocus
                End If
        End If
End Sub
