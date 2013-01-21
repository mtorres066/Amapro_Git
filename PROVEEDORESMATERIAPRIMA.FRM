VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ProveedoresMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento De Proveedores Materia Prima"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "ProveedoresMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataProveedores 
      Caption         =   "Proveedores"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProveedoresMateriaPrima"
      Top             =   5040
      Width           =   9135
   End
   Begin TabDlg.SSTab TabDepartamentos 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "ProveedoresMateriaPrima.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProveedores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "ProveedoresMateriaPrima.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridProveedores"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "ProveedoresMateriaPrima.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridProveedores 
         Bindings        =   "ProveedoresMateriaPrima.frx":1488
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "ProveedoresMateriaPrima.frx":14A6
         TabIndex        =   16
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
         Height          =   4575
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   8775
         Begin VB.TextBox TxtBusqueda 
            Height          =   285
            Left            =   5760
            TabIndex        =   19
            Top             =   720
            Width           =   2775
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Descripcion"
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
            Picture         =   "ProveedoresMateriaPrima.frx":2A4E
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   975
            Index           =   0
            Left            =   360
            Picture         =   "ProveedoresMateriaPrima.frx":2E90
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
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
            Left            =   3600
            TabIndex        =   30
            Top             =   720
            Width           =   2055
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   6000
            TabIndex        =   21
            Top             =   3120
            Width           =   2535
            Caption         =   "Actualizar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "ProveedoresMateriaPrima.frx":319A
            Accelerator     =   84
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
            TabIndex        =   20
            Top             =   2400
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "ProveedoresMateriaPrima.frx":34B4
            Accelerator     =   83
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameProveedores 
         Caption         =   "Datos del Proveedor"
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
         Height          =   3855
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "DiasDeEntrega"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   7
            Left            =   2280
            TabIndex        =   9
            Top             =   3360
            Width           =   840
         End
         Begin VB.ComboBox CboTipPro 
            BackColor       =   &H008080FF&
            DataField       =   "TipoDeProveedor"
            DataSource      =   "DataProveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "ProveedoresMateriaPrima.frx":37CE
            Left            =   2280
            List            =   "ProveedoresMateriaPrima.frx":37D8
            TabIndex        =   1
            Text            =   "LOCAL"
            ToolTipText     =   "Seleccione Tipo De Proveedor"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Nit"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   6
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   7
            Top             =   2640
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Encargado"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   5
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3000
            Width           =   6240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Fax"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   4
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   6
            Top             =   2235
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Telefono"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   3
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1860
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Direccion"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   2
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1485
            Width           =   6240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Proveedor"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1095
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CodigoProveedor"
            DataSource      =   "DataProveedores"
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   2
            Top             =   720
            Width           =   1980
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Dias De Entrega"
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   33
            Top             =   3360
            Width           =   1170
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tipo De Proveedor"
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
            Index           =   7
            Left            =   480
            TabIndex        =   32
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nit"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   31
            Top             =   2640
            Width           =   195
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Encargado"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   28
            Top             =   3000
            Width           =   780
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   27
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   26
            Top             =   1905
            Width           =   630
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   25
            Top             =   1530
            Width           =   675
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   24
            Top             =   1140
            Width           =   555
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   23
            Top             =   765
            Width           =   600
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "2355;1085"
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
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
      Caption         =   "Editar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "ProveedoresMateriaPrima.frx":37F2
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
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "ProveedoresMateriaPrima.frx":3D34
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
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "2566;1085"
      Picture         =   "ProveedoresMateriaPrima.frx":4276
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
      TabIndex        =   14
      Top             =   5640
      Width           =   1335
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "ProveedoresMateriaPrima.frx":47B8
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
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "ProveedoresMateriaPrima.frx":4CFA
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "ProveedoresMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer



Private Sub CboTipPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                SendKeys "{tab}"
        End If
End Sub

Private Sub CboTipPro_LostFocus()
    CboTipPro.Text = UCase(CboTipPro.Text)
End Sub

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataProveedores.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                CboTipPro.SetFocus
                CboTipPro.Text = "LOCAL"
        
        'EDITAR
        ElseIf Index = 1 Then
                        .Edit
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                CboTipPro.SetFocus
        
        'GRABAR
        ElseIf Index = 2 Then
                    If (CboTipPro.Text <> "LOCAL") And (CboTipPro.Text <> "INTERNACIONAL") Then
                            MsgBox "Tipo De Proveedore Incorrecto", vbOKOnly + vbInformation, "Informacion"
                            CboTipPro.SetFocus
                            Exit Sub
                    End If
        
                .Update
                
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 3 Then
                'DataProveedores.Refresh
                .CancelUpdate
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number > 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataProveedores.RecordSource = ("Select * From ProveedoresMateriaPrima where CodigoProveedor like '" & TxtBusqueda.Text & "*'")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataProveedores.RecordSource = ("Select * From ProveedoresMateriaPrima where Proveedor like '" & TxtBusqueda.Text & "*'")
                    End If
                    DataProveedores.Refresh
                    DGridProveedores.Refresh
                    TabDepartamentos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataProveedores.RecordSource = "Select * From ProveedoresMateriaPrima"
                    DataProveedores.Refresh
                    DGridProveedores.Refresh
                    TabDepartamentos.Tab = 1
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
         FrameProveedores.Enabled = True
         DataProveedores.Visible = False
         DGridProveedores.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameProveedores.Enabled = False
         DataProveedores.Visible = True
         DGridProveedores.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub Form_Load()
    DataProveedores.DatabaseName = BasedeDatos
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Codigo"
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Descripcion"
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
End Sub
