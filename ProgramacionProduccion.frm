VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ProgramacionProduccion 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programacion De la Produccion"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "ProgramacionProduccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   7215
      Left            =   0
      TabIndex        =   35
      Top             =   0
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
         TabIndex        =   40
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   42
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   7560
         Picture         =   "ProgramacionProduccion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "ProgramacionProduccion.frx":0D0C
         Height          =   6015
         Left            =   120
         OleObjectBlob   =   "ProgramacionProduccion.frx":0D28
         TabIndex        =   39
         Top             =   1080
         Width           =   8175
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
   End
   Begin TabDlg.SSTab TabEmpleados 
      Height          =   5775
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "ProgramacionProduccion.frx":1703
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProgramacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "ProgramacionProduccion.frx":1A1D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridProgramacion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "ProgramacionProduccion.frx":1E6F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtBusqueda"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "DtpFecFin"
      Tab(2).Control(3)=   "DtpFecIni"
      Tab(2).Control(4)=   "CmdBuscar(1)"
      Tab(2).Control(5)=   "CmdBuscar(0)"
      Tab(2).Control(6)=   "LblEti"
      Tab(2).Control(7)=   "Label3"
      Tab(2).Control(8)=   "Label1"
      Tab(2).ControlCount=   9
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68760
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones De Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Ficha Tecnica"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Turno"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Linea"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -68760
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   64880641
         CurrentDate     =   37588
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -70680
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64880643
         CurrentDate     =   37588
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "ProgramacionProduccion.frx":22C1
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "ProgramacionProduccion.frx":25CB
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Frame FrameProgramacion 
         Caption         =   "Datos Del La Programacion"
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
         Height          =   2655
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   8175
         Begin MSMask.MaskEdBox MskCan 
            DataField       =   "Cantidad"
            DataSource      =   "DataProgramacion"
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Usuario"
            DataSource      =   "DataProgramacion"
            Height          =   285
            Index           =   2
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataProgramacion"
            Height          =   285
            Index           =   3
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataProgramacion"
            Height          =   285
            Left            =   1200
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
            DataSource      =   "DataProgramacion"
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Turno"
            DataSource      =   "DataProgramacion"
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label LblFicTec 
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
            Left            =   2760
            TabIndex        =   31
            Top             =   1440
            Width           =   5295
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
            Left            =   2760
            TabIndex        =   30
            Top             =   720
            Width           =   5295
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   420
         End
      End
      Begin MSDBGrid.DBGrid DBGridProgramacion 
         Bindings        =   "ProgramacionProduccion.frx":2A0D
         Height          =   4905
         Left            =   -74880
         OleObjectBlob   =   "ProgramacionProduccion.frx":2A2C
         TabIndex        =   12
         Top             =   720
         Width           =   8145
      End
      Begin VB.Label LblEti 
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
         TabIndex        =   45
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   -69240
         TabIndex        =   33
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   -71160
         TabIndex        =   32
         Top             =   1680
         Width           =   465
      End
   End
   Begin VB.Data DataProgramacion 
      BackColor       =   &H80000014&
      Caption         =   "Programacion De La Produccion"
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
      RecordSource    =   "ProgramacionProduccion"
      Top             =   6840
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "ProgramacionProduccion.frx":3AC2
      Picture         =   "ProgramacionProduccion.frx":3F04
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1320
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "ProgramacionProduccion.frx":5F76
      Picture         =   "ProgramacionProduccion.frx":63B8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "ProgramacionProduccion.frx":68EA
      Picture         =   "ProgramacionProduccion.frx":6D2C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "ProgramacionProduccion.frx":725E
      Picture         =   "ProgramacionProduccion.frx":76A0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "ProgramacionProduccion.frx":7BD2
      Picture         =   "ProgramacionProduccion.frx":8014
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "ProgramacionProduccion.frx":8546
      Picture         =   "ProgramacionProduccion.frx":8988
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1200
   End
End
Attribute VB_Name = "ProgramacionProduccion"
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
Dim BFichaTecnica As Boolean

Dim RBuscaLineas As Recordset
Dim RBuscaFichaTecnica As Recordset



Sub botones()
    If Bandera = True Then
         FrameProgramacion.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         DataProgramacion.Visible = False
         DBGridProgramacion.Visible = False
    Else
         FrameProgramacion.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataProgramacion.Visible = True
         DBGridProgramacion.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataProgramacion.Recordset
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
                    Txttexto.Item(1).SetFocus
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
                    Txttexto.Item(1).SetFocus
                    Txttexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                   
                     'GRABA EL REGISTRO
                     .Update
                      'SI ES CUALQUIER OTRO ERROR
                     If Err <> 3022 And Err <> 0 Then
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
                        DataProgramacion.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataProgramacion.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataProgramacion.Recordset.EOF Then
                        DataProgramacion.Recordset.MoveLast
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
    With DataProgramacion
        'SELECCIONAR DATOS
        If Index = 0 Then
                If OptOpcion.Item(0).Value = True Then
                    .RecordSource = ("Select * from ProgramacionProduccion where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "#")
                ElseIf OptOpcion.Item(1).Value = True Then
                    .RecordSource = ("Select * from ProgramacionProduccion where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "#" & " And Linea = '" & TxtBusqueda.Text & "'")
                ElseIf OptOpcion.Item(2).Value = True Then
                    .RecordSource = ("Select * from ProgramacionProduccion where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "#" & " And Turno = '" & TxtBusqueda.Text & "'")
                ElseIf OptOpcion.Item(3).Value = True Then
                    .RecordSource = ("Select * from ProgramacionProduccion where Fecha >= #" & DtpFecIni.Value & "# And Fecha <= #" & DtpFecFin.Value & "#" & " And FichaTecnica = '" & TxtBusqueda.Text & "'")
                End If
                .Refresh
                DBGridProgramacion.Refresh
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From ProgramacionProduccion"
                .Refresh
                DBGridProgramacion.Refresh
        End If
    End With
        TabEmpleados.Tab = 1
End Sub


Private Sub Command1_Click()
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
        If BTurno = True Then
            Txttexto.Item(0).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(0).SetFocus
        ElseIf BLinea = True Then
            Txttexto.Item(1).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(1).SetFocus
        ElseIf BFichaTecnica = True Then
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
        ElseIf BFichaTecnica = True Then
            Txttexto.Item(3).Text = DBGridConsultas.Columns(0).Text
            Txttexto.Item(3).SetFocus
        End If
        FrameConsultas.Visible = False
End Sub



Private Sub DBGridProgramacion_HeadClick(ByVal ColIndex As Integer)
DataProgramacion.RecordSource = ("Select * from ProgramacionProduccion order by " & DBGridProgramacion.Columns(ColIndex).DataField)
    DataProgramacion.Refresh
    DBGridProgramacion.Refresh

End Sub

Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        DataProgramacion.Connect = GConnect
        DataConsultas.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        DataProgramacion.DatabaseName = BasedeDatos
        DataConsultas.DatabaseName = BasedeDatos

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

Private Sub MskFec_GotFocus()
        MskFec.SelStart = 0
        MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub MskFec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub OptOpcion_Click(Index As Integer)
    If Index = 0 Then
        TxtBusqueda.Visible = False
    Else
        If Index = 1 Then
            LblEti.Caption = "Linea"
        ElseIf Index = 2 Then
            LblEti.Caption = "Turno"
        ElseIf Index = 3 Then
            LblEti.Caption = "Ficha Tecnica"
        End If
        TxtBusqueda.Visible = True
        TxtBusqueda.SetFocus
    End If
End Sub

Private Sub tabempleados_Click(PreviousTab As Integer)
        If TabEmpleados.Tab = 2 Then
            DtpFecIni.Value = Date
            DtpFecFin.Value = Date
        End If
End Sub

Private Sub TxtConsultas_Change()
    'FICHA TECNICA
    If BFichaTecnica = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '" & TxtConsultas.Text & "*' Order By Descrip"
            Else
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '*" & TxtConsultas.Text & "*' Order By Descrip"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '" & TxtConsultas.Text & "*' Order By Descrip"
            Else
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '*" & TxtConsultas.Text & "*' Order By Descrip"
            End If
        End If
    'LINEA
    ElseIf BLinea = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas Where Descrip Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas Where Descrip Like '*" & TxtConsultas.Text & "*'"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas Where Linea Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas Where Linea Like '*" & TxtConsultas.Text & "*'"
            End If
        End If
    End If
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"

End Sub

Private Sub TxtConsultas_GotFocus()
        TxtConsultas.SelStart = 0
        TxtConsultas.SelLength = Len(TxtConsultas.Text)
End Sub

Private Sub TxtConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
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
            Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & Txttexto.Item(3).Text & "'")
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFicTec.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFicTec.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 0 Then
                BTurno = True
                BLinea = False
                BFichaTecnica = False
                DataConsultas.RecordSource = "Select * From Turnos"
        ElseIf Index = 1 Then
                BTurno = False
                BLinea = True
                BFichaTecnica = False
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
        ElseIf Index = 3 Then
                BTurno = False
                BLinea = False
                BFichaTecnica = True
                DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
        End If
        
        If Index = 0 Or Index = 1 Or Index = 3 Then
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                DBGridConsultas.Columns(1).Width = "4000"
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
                        BFichaTecnica = False
                        DataConsultas.RecordSource = "Select * From Turnos"
                ElseIf Index = 1 Then
                        BTurno = False
                        BLinea = True
                        BFichaTecnica = False
                        DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
                ElseIf Index = 3 Then
                        BTurno = False
                        BLinea = False
                        BFichaTecnica = True
                        DataConsultas.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
                End If
                
                If Index = 0 Or Index = 1 Or Index = 3 Then
                        DataConsultas.Refresh
                        DBGridConsultas.Refresh
                        DBGridConsultas.Columns(1).Width = "4000"
                        FrameConsultas.Visible = True
                        TxtConsultas.SetFocus
                End If
        End If
End Sub
