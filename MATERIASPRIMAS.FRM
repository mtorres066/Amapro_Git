VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MateriasPrimas 
   Caption         =   "Materias Primas"
   ClientHeight    =   6750
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   12015
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   5295
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Data DataConsultas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Escuintla\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Lineas"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10680
         Picture         =   "MateriasPrimas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "MateriasPrimas.frx":030A
         Height          =   4935
         Left            =   120
         OleObjectBlob   =   "MateriasPrimas.frx":0326
         TabIndex        =   19
         Top             =   240
         Width           =   10335
      End
   End
   Begin TabDlg.SSTab TabMateriasPrimas 
      Height          =   6375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "MateriasPrimas.frx":0D01
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CmdBotones(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdBotones(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdBotones(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBotones(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBotones(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameMateriasPrimas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdBotones(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "MateriasPrimas.frx":101B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridMateriasPrimas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "MateriasPrimas.frx":146D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusqueda"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Fotografia"
      TabPicture(3)   =   "MateriasPrimas.frx":18BF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ImageFoto"
      Tab(3).ControlCount=   1
      Begin VB.Frame FrameBusqueda 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   32
         Top             =   720
         Width           =   11655
         Begin VB.OptionButton OptLinFec 
            Caption         =   "Linea Y Fecha"
            Height          =   855
            Left            =   840
            Picture         =   "MateriasPrimas.frx":2199
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   840
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8760
            TabIndex        =   36
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   8760
            TabIndex        =   35
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Datos"
            Height          =   735
            Index           =   0
            Left            =   8280
            Picture         =   "MateriasPrimas.frx":24A3
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2520
            Width           =   2655
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Todos Los Datos"
            Height          =   735
            Index           =   1
            Left            =   8280
            Picture         =   "MateriasPrimas.frx":28E5
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Linea"
            Height          =   255
            Index           =   0
            Left            =   6720
            TabIndex        =   39
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   38
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   825
         Index           =   3
         Left            =   6120
         MouseIcon       =   "MateriasPrimas.frx":2BEF
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":3031
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1900
      End
      Begin VB.Frame FrameMateriasPrimas 
         Caption         =   "Datos Materias Primas"
         Enabled         =   0   'False
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   11655
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataMateriasPrimas"
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
            Index           =   10
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   41
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Usuario"
            DataSource      =   "DataMateriasPrimas"
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
            Index           =   9
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "NylonStrech"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   8
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   10
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataMateriasPrimas"
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
            Index           =   0
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Fecha"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Hojalata"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Fondo"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   3
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "AlambreCobre"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   6
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BarnizLiquido"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BarnizPolvo"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   6
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   8
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "SelloSolvente"
            DataSource      =   "DataMateriasPrimas"
            Height          =   285
            Index           =   7
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   9
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label LblFicha 
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
            Left            =   3480
            TabIndex        =   43
            Top             =   960
            Width           =   6735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ficha Tecnica"
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
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Usuario"
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
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LblLinea 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   31
            Top             =   600
            Width           =   6015
         End
         Begin VB.Label lblLabels 
            Caption         =   "Linea"
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
            Index           =   18
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hojalata"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fondo"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   27
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Alambre Cobre"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   26
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Barniz Liquido"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   25
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Barniz Polvo"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   24
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Sello Solvente"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   23
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Nylon Strech"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   22
            Top             =   3840
            Width           =   1215
         End
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Agregar"
         Height          =   825
         Index           =   0
         Left            =   360
         MouseIcon       =   "MateriasPrimas.frx":3563
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":39A5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Editar"
         Height          =   825
         Index           =   1
         Left            =   2280
         MouseIcon       =   "MateriasPrimas.frx":3ED7
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":4319
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   825
         Index           =   2
         Left            =   4200
         MouseIcon       =   "MateriasPrimas.frx":484B
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":4C8D
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "B&orrar"
         Height          =   825
         Index           =   4
         Left            =   8040
         MouseIcon       =   "MateriasPrimas.frx":51BF
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":5601
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Salida"
         Height          =   825
         Index           =   5
         Left            =   9960
         MouseIcon       =   "MateriasPrimas.frx":5B33
         MousePointer    =   99  'Custom
         Picture         =   "MateriasPrimas.frx":5F75
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   " "
         Top             =   5280
         Width           =   1785
      End
      Begin MSDBGrid.DBGrid DBGridMateriasPrimas 
         Bindings        =   "MateriasPrimas.frx":627F
         Height          =   4935
         Left            =   -74880
         OleObjectBlob   =   "MateriasPrimas.frx":62A0
         TabIndex        =   17
         Top             =   720
         Width           =   11535
      End
      Begin VB.Image ImageFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   5415
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   720
         Width           =   7815
      End
   End
   Begin VB.Data DataMateriasPrimas 
      Align           =   2  'Align Bottom
      Caption         =   "Materias Primas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Escuintla\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MateriasPrimas"
      Top             =   6405
      Width           =   12015
   End
End
Attribute VB_Name = "MateriasPrimas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim RLineas As Recordset

Dim RUsuarios As Recordset
Dim REquipos As Recordset
Dim RLineaActiva As Recordset
Dim RFicha As Recordset
Dim RBuscaFoto As Recordset


Sub botones()
    If Bandera = True Then
         FrameMateriasPrimas.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         DataMateriasPrimas.Visible = False
         DBGridMateriasPrimas.Visible = False
         FrameBusqueda.Visible = False
    Else
         FrameMateriasPrimas.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataMateriasPrimas.Visible = True
         DBGridMateriasPrimas.Visible = True
         FrameBusqueda.Visible = True
    End If
End Sub



Private Sub CmdBotones_Click(Index As Integer)
    

                'AGREGAR
                If Index = 0 Then
                                    On Error Resume Next
                                    DataMateriasPrimas.Recordset.AddNew
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                            Bandera = True
                                            botones
                                            TxtTexto.Item(9).Text = GUsuario
                                            TxtTexto.Item(1).Text = Format(Date, "dd/mm/yyyy")
                                            TxtTexto.Item(2).SetFocus
                                            
                                        '----------------------------------------------------------------PROCESO PARA GENERAR LINEA DEL USUARIO DE ACUERDO AL EQUIPO QUE PERTENECE
                                        
                                                        'BUSCA EL EQUIPO AL QUE PERTENECE EL USUARIO
                                                        Set RUsuarios = Db.OpenRecordset("Select Equipo From Usuarios Where Usuario = '" & TxtTexto.Item(9).Text & "'")
                                                        If RUsuarios.RecordCount > 0 Then
                                                                    'BUSCA LAS LINEAS QUE ESTAN ASIGNADAS A EL EQUIPO DEL USUARIO
                                                                    Set REquipos = Db.OpenRecordset("Select Linea From Equipos Where Equipo = '" & RUsuarios(0) & "'")
                                                                            If REquipos.RecordCount > 0 Then
                                                                                    Do Until REquipos.EOF
                                                                                                    Set RLineaActiva = Db.OpenRecordset("Select Linea, Esp_Tec From Lineas Where Linea = '" & REquipos(0) & "' and Activa = -1")
                                                                                                            'SI ENCUENTRA LA LINEA ACTIVA QUE TIENE ASIGNADA EL USUARIO LA ASIGNA
                                                                                                            If RLineaActiva.RecordCount > 0 Then
                                                                                                                TxtTexto.Item(0).Text = RLineaActiva(0)
                                                                                                                TxtTexto.Item(10).Text = RLineaActiva(1)
                                                                                                                Exit Sub
                                                                                                            End If
                                                                                            REquipos.MoveNext
                                                                                    Loop
                                                                            Else
                                                                                MsgBox "No Hay Lineas Asignadas Para El Equipo Al Que Pertenece El Usuario", vbInformation, "Informacion"
                                                                            End If
                                                        Else
                                                                MsgBox "Usuario No Existe", vbInformation
                                                        End If
                                        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                               
                                    End If
                                    
                'EDITAR
                ElseIf Index = 1 Then
                                    On Error Resume Next
                                    DataMateriasPrimas.Recordset.Edit
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                            Bandera = True
                                            botones
                                            TxtTexto.Item(0).SetFocus
                                    End If
                                    
                'GRABAR
                ElseIf Index = 2 Then
                                    On Error Resume Next
                
                                    DataMateriasPrimas.Recordset.Update
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                       Bandera = False
                                       botones
                                   End If
                'CANCELAR
                ElseIf Index = 3 Then
                                    On Error Resume Next
                                    DataMateriasPrimas.Recordset.CancelUpdate
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                        Bandera = False
                                        botones
                                    End If
                'BORRAR
                ElseIf Index = 4 Then
                                    On Error Resume Next
                                            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                                
                                            If mensaje = vbOK Then
                                                DataMateriasPrimas.Recordset.Delete
                                                DataMateriasPrimas.Recordset.MoveLast
                                            End If
                                  
                                            If DataMateriasPrimas.Recordset.EOF Then
                                                DataMateriasPrimas.Recordset.MoveLast
                                                If Err = 3021 Then
                                                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                                                End If
                                            End If
                                           
                                           If Err <> 0 Then
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                           End If
                                            
                'SALIDA
                Else
                                        Unload Me
                End If
    
    
End Sub

Private Sub CmdBusqueda_Click(Index As Integer)
MousePointer = 11
If Index = 0 Then
    TxtBusqueda.Item(1).Text = Format(TxtBusqueda.Item(1).Text, "dd/mm/yyyy")
    If Not IsDate(TxtBusqueda.Item(1).Text) Then
            MsgBox "Fecha Incorrecta", vbInformation, "Informacion"
            Exit Sub
    End If

    DataMateriasPrimas.RecordSource = "Select * from MateriasPrimas Where Linea = '" & TxtBusqueda.Item(0).Text & "' And Fecha = #" & Format(TxtBusqueda.Item(1).Text, "mm/dd/yyyy") & "#"
    DataMateriasPrimas.Refresh
    DBGridMateriasPrimas.Refresh
End If

If Index = 1 Then
    DataMateriasPrimas.RecordSource = "Select * from MateriasPrimas"
    DataMateriasPrimas.Refresh
    DBGridMateriasPrimas.Refresh
End If

MousePointer = 0
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
                    TxtTexto.Item(0).Text = DBGridConsultas.Columns(0).Text
                    FrameConsultas.Visible = False
                    TxtTexto.Item(0).SetFocus
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then
                    TxtTexto.Item(0).Text = DBGridConsultas.Columns(0).Text
                    FrameConsultas.Visible = False
                    TxtTexto.Item(0).SetFocus
        End If
End Sub

Private Sub dbgridmateriasprimas_HeadClick(ByVal ColIndex As Integer)
    DataMateriasPrimas.RecordSource = ("Select * from MateriasPrimas order by " & DBGridMateriasPrimas.Columns(ColIndex).DataField)
    DataMateriasPrimas.Refresh
    DBGridMateriasPrimas.Refresh
    
End Sub

Private Sub Form_Load()
    DataMateriasPrimas.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub



Private Sub Tab_Click(PreviousTab As Integer)

        
        
End Sub

Private Sub Tab_DblClick()

End Sub

Private Sub TabMateriasPrimas_Click(PreviousTab As Integer)
On Error Resume Next
        
        If TabMateriasPrimas.Tab = 2 Then
                If FrameBusqueda.Visible = True Then
                    TxtBusqueda.Item(0).SetFocus
                End If
        End If
        


If TabMateriasPrimas.Tab = 3 Then

    Set RBuscaFoto = Db.OpenRecordset("Select Foto1 From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(10).Text & "'")
    If RBuscaFoto.RecordCount > 0 Then
                    
                   ImageFoto.Picture = LoadPicture(App.Path & "\Fotos\" & RBuscaFoto!Foto1)
                    
                    If Err.Number > 0 Then
                              ImageFoto.Picture = LoadPicture("")
                    End If
    End If
    
End If

       
        
End Sub

Private Sub TxtBusqueda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
'LINEAS
If Index = 0 Then
        Set RLineas = Db.OpenRecordset("select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
        If RLineas.RecordCount > 0 Then
                LblLinea.Caption = RLineas(0)
        Else
                LblLinea.Caption = ""
        End If
'FICHA TECNICA
ElseIf Index = 10 Then
        Set RFicha = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(10).Text & "'")
        If RFicha.RecordCount > 0 Then
                LblFicha.Caption = RFicha(0)
        Else
                LblFicha.Caption = ""
        End If
        

End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
If Index = 0 Then
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
    End If
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
    If Index = 0 Then
            If KeyAscii = 43 Then
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
            End If
    End If
End Sub

