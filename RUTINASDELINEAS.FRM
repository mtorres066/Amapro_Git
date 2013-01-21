VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form RutinasDeLineas 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activacion De Rutinas Para Las Lineas De Produccion"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "RutinasDeLineas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda de Datos"
      Height          =   7215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton CmdSale 
         Height          =   855
         Left            =   8160
         Picture         =   "RutinasDeLineas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Sale de Busqueda"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "RutinasDeLineas.frx":074C
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "RutinasDeLineas.frx":0767
         TabIndex        =   10
         ToolTipText     =   "Doble click o signo '+' para ayuda"
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Data DataRutinasLineas 
      Caption         =   "Rutinas De Lineas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Erick\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RutinasDeLineas"
      Top             =   8040
      Width           =   8895
   End
   Begin TabDlg.SSTab TabFondos 
      Height          =   7215
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual y General"
      TabPicture(0)   =   "RutinasDeLineas.frx":1141
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGridRutinasLineas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameRutinasLineas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Busqueda De Datos"
      TabPicture(1)   =   "RutinasDeLineas.frx":145B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameBusquedadeDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
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
         TabIndex        =   18
         Top             =   1680
         Width           =   8775
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Codigo Linea"
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
            Left            =   360
            Picture         =   "RutinasDeLineas.frx":18AD
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Codigo Rutina"
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
            Left            =   1800
            Picture         =   "RutinasDeLineas.frx":1BB7
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   19
            Top             =   1800
            Width           =   2535
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   6
            Left            =   6000
            TabIndex        =   20
            Top             =   2400
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "RutinasDeLineas.frx":1EC1
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   6000
            TabIndex        =   21
            Top             =   3120
            Width           =   2535
            Caption         =   "Seleccionar Todos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "RutinasDeLineas.frx":21DB
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
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
            TabIndex        =   24
            Top             =   1800
            Width           =   2055
         End
      End
      Begin VB.Frame FrameRutinasLineas 
         Caption         =   "Datos De La Rutina Y Linea"
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
         Height          =   1095
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CodigoRutina"
            DataSource      =   "DataRutinasLineas"
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CodigoLinea"
            DataSource      =   "DataRutinasLineas"
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label LblRutina 
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
            Left            =   3120
            TabIndex        =   16
            Top             =   720
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
            Left            =   3120
            TabIndex        =   15
            Top             =   360
            Width           =   5295
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo De Rutina"
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
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo De Linea"
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
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1425
         End
      End
      Begin MSDBGrid.DBGrid DGridRutinasLineas 
         Bindings        =   "RutinasDeLineas.frx":24F5
         Height          =   5175
         Left            =   240
         OleObjectBlob   =   "RutinasDeLineas.frx":2515
         TabIndex        =   17
         Top             =   1920
         Width           =   8655
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   1335
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "RutinasDeLineas.frx":2F1B
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
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
      Caption         =   "Editar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "RutinasDeLineas.frx":345D
      Accelerator     =   69
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
      Left            =   3000
      TabIndex        =   5
      Top             =   7320
      Width           =   1335
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "RutinasDeLineas.frx":399F
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
      Left            =   4440
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "2566;1085"
      Picture         =   "RutinasDeLineas.frx":3EE1
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
      Left            =   6000
      TabIndex        =   7
      Top             =   7320
      Width           =   1335
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "RutinasDeLineas.frx":4423
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
      Left            =   7440
      TabIndex        =   8
      Top             =   7320
      Width           =   1335
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "RutinasDeLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim BLinea As Boolean
Dim BRutina As Boolean

Dim RBuscaLinea As Recordset
Dim RBuscaRutina As Recordset

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataRutinasLineas.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                        TxtTexto.Item(0).SetFocus
        'EDITAR
        ElseIf Index = 1 Then
                        .Edit
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                        TxtTexto.Item(0).SetFocus
        'GRABAR
        ElseIf Index = 2 Then
                            
                        'REVISA SI EXISTE LA LINEA
                        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
                            If RBuscaLinea.RecordCount > 0 Then
                            Else
                                MsgBox "Codigo De Linea No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(0).SetFocus
                                Exit Sub
                            End If
                            
                       'REVISA SI EXISTE LA LINEA
                        Set RBuscaRutina = Db.OpenRecordset("Select Descrip From Rutinas Where Rutina = '" & TxtTexto.Item(1).Text & "'")
                            If RBuscaRutina.RecordCount > 0 Then
                            Else
                                MsgBox "Codigo De Rutina No Existe", vbOKOnly + vbInformation, "Informacion"
                                TxtTexto.Item(1).SetFocus
                                Exit Sub
                            End If
                    
                
                .Update
                        'SI SE DUPLICA LA LLAVE
                        If Err.Number = 3022 Then
                                MsgBox "Codigo De Linea Con Codigo De Rutina Ya Existen", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                        'POR CUALQUIER OTRO ERROR
                        ElseIf Err.Number <> 0 And Err.Number <> 3022 And Err.Number <> 3201 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 3 Then
                .CancelUpdate
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
        
                VMensaje = MsgBox("Esta Seguro De Borrar El Registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataRutinasLineas.RecordSource = ("Select * From RutinasDeLineas Where CodigoLinea like '" & TxtBusqueda.Text & "*'")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataRutinasLineas.RecordSource = ("Select * From RutinasDeLineas where CodigoRutina like '" & TxtBusqueda.Text & "*'")
                    End If
                    DataRutinasLineas.Refresh
                    DGridRutinasLineas.Refresh
                    TabFondos.Tab = 0
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataRutinasLineas.RecordSource = "Select * From RutinasDeLineas"
                    DataRutinasLineas.Refresh
                    DGridRutinasLineas.Refresh
                    TabFondos.Tab = 0
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameRutinasLineas.Enabled = True
         DataRutinasLineas.Visible = False
         DGridRutinasLineas.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameRutinasLineas.Enabled = False
         DataRutinasLineas.Visible = True
         DGridRutinasLineas.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
    If BLinea = True Then
        TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
        TxtTexto.Item(0).SetFocus
    ElseIf BRutina = True Then
        TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
        TxtTexto.Item(1).SetFocus
    End If
        FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        If BLinea = True Then
            TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(0).SetFocus
        ElseIf BRutina = True Then
            TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(1).SetFocus
        End If
            FrameBusqueda.Visible = False
    End If
End Sub

Private Sub dgridrutinaslineas_HeadClick(ByVal ColIndex As Integer)
        DataRutinasLineas.RecordSource = "Select * From RutinasDeLineas Order by " & DGridRutinasLineas.Columns(ColIndex).DataField
        DataRutinasLineas.Refresh
        DGridRutinasLineas.Refresh
End Sub

Private Sub Form_Load()
    DataRutinasLineas.DatabaseName = BasedeDatos
    DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Codigo De Linea"
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Codigo De Rutina"
    End If
            TxtBusqueda.SetFocus
    
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
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        ElseIf Index = 1 Then
            Set RBuscaRutina = Db.OpenRecordset("Select Descrip From Rutinas Where Rutina = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaRutina.RecordCount > 0 Then
                    LblRutina.Caption = RBuscaRutina!Descrip
                Else
                    LblRutina.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    If Index = 0 Then
        BLinea = True
        BRutina = False
        DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas"
    ElseIf Index = 1 Then
        BLinea = False
        BRutina = True
        DataBusqueda.RecordSource = "Select Rutina, Descrip, Cabezal From Rutinas"
    End If
        DataBusqueda.Refresh
        DBGridBusqueda.Refresh
        DBGridBusqueda.Columns(1).Width = "4000"
        FrameBusqueda.Visible = True
        DBGridBusqueda.SetFocus
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
            If Index = 0 Then
                BLinea = True
                BRutina = False
                DataBusqueda.RecordSource = "Select Linea, Descrip From Lineas"
            ElseIf Index = 1 Then
                BLinea = False
                BRutina = True
                DataBusqueda.RecordSource = "Select Rutina, Descrip, Cabezal From Rutinas"
            End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                DBGridBusqueda.SetFocus
        End If
End Sub
