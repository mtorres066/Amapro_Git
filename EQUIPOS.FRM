VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Equipos 
   BackColor       =   &H000000FF&
   Caption         =   "Equipos"
   ClientHeight    =   4125
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   12015
   ControlBox      =   0   'False
   Icon            =   "Equipos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabMateriasPrimas 
      Height          =   3615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Equipos.frx":0442
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
      Tab(0).Control(5)=   "FrameEquipos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdBotones(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Equipos.frx":075C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridEquipos"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   825
         Index           =   3
         Left            =   6000
         MouseIcon       =   "Equipos.frx":0BAE
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1900
      End
      Begin VB.Frame FrameEquipos 
         Caption         =   "Datos del Equipo"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   240
         TabIndex        =   0
         Top             =   960
         Width           =   11655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Equipo"
            DataSource      =   "DataEquipos"
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
            MaxLength       =   50
            TabIndex        =   1
            Top             =   360
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataEquipos"
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   2
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Equipo"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Linea"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Agregar"
         Height          =   825
         Index           =   0
         Left            =   240
         MouseIcon       =   "Equipos.frx":1522
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":1964
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Editar"
         Height          =   825
         Index           =   1
         Left            =   2160
         MouseIcon       =   "Equipos.frx":1E96
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":22D8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   825
         Index           =   2
         Left            =   4080
         MouseIcon       =   "Equipos.frx":280A
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":2C4C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "B&orrar"
         Height          =   825
         Index           =   4
         Left            =   7920
         MouseIcon       =   "Equipos.frx":317E
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":35C0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1900
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Salida"
         Height          =   825
         Index           =   5
         Left            =   9840
         MouseIcon       =   "Equipos.frx":3AF2
         MousePointer    =   99  'Custom
         Picture         =   "Equipos.frx":3F34
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1785
      End
      Begin MSDBGrid.DBGrid DBGridEquipos 
         Bindings        =   "Equipos.frx":423E
         Height          =   2775
         Left            =   -74880
         OleObjectBlob   =   "Equipos.frx":4258
         TabIndex        =   9
         Top             =   720
         Width           =   11535
      End
   End
   Begin VB.Data DataEquipos 
      Caption         =   "Equipos"
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
      RecordSource    =   "Equipos"
      Top             =   3720
      Width           =   12015
   End
End
Attribute VB_Name = "Equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String



Sub botones()
    If Bandera = True Then
         FrameEquipos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         DataEquipos.Visible = False
         DBGridEquipos.Visible = False
    Else
         FrameEquipos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataEquipos.Visible = True
         DBGridEquipos.Visible = True
    End If
End Sub



Private Sub CmdBotones_Click(Index As Integer)
    

                'AGREGAR
                If Index = 0 Then
                                    On Error Resume Next
                                    DataEquipos.Recordset.AddNew
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                            Bandera = True
                                            botones
                                            TxtTexto.Item(0).SetFocus
                                    End If
                                    
                'EDITAR
                ElseIf Index = 1 Then
                                    On Error Resume Next
                                    DataEquipos.Recordset.Edit
                                    
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
                
                                    DataEquipos.Recordset.Update
                                    
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                    Else
                                       Bandera = False
                                       botones
                                       CmdBotones.Item(0).SetFocus
                                   End If
                'CANCELAR
                ElseIf Index = 3 Then
                                    On Error Resume Next
                                    DataEquipos.Recordset.CancelUpdate
                                    
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
                                                DataEquipos.Recordset.Delete
                                                DataEquipos.Recordset.MoveLast
                                            End If
                                  
                                            If DataEquipos.Recordset.EOF Then
                                                DataEquipos.Recordset.MoveLast
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


Private Sub dbgridequipos_HeadClick(ByVal ColIndex As Integer)
    DataEquipos.RecordSource = ("Select * from Equipos order by " & DBGridEquipos.Columns(ColIndex).DataField)
    DataEquipos.Refresh
    DBGridEquipos.Refresh
    
End Sub

Private Sub Form_Load()
    DataEquipos.DatabaseName = BasedeDatos
    
End Sub



Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
End Sub
