VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Platinas 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Tecnica De HOJALATAS"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "Platinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda de Datos"
      Height          =   5535
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton CmdSale 
         Height          =   855
         Left            =   8160
         Picture         =   "Platinas.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Bindings        =   "Platinas.frx":130C
         Height          =   5175
         Left            =   120
         OleObjectBlob   =   "Platinas.frx":1327
         TabIndex        =   23
         ToolTipText     =   "Doble click o signo '+' para ayuda"
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Data DataHojalatas 
      Caption         =   "Hojalatas"
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
      RecordSource    =   "Platinas"
      Top             =   5640
      Width           =   8895
   End
   Begin TabDlg.SSTab TabHojalatas 
      Height          =   4815
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Platinas.frx":1D01
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameHojalatas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Platinas.frx":201B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridHojalatas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "Platinas.frx":246D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridHojalatas 
         Bindings        =   "Platinas.frx":28BF
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "Platinas.frx":28DB
         TabIndex        =   14
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
         TabIndex        =   15
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
            Picture         =   "Platinas.frx":3961
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Picture         =   "Platinas.frx":3DA3
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   18
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
            TabIndex        =   20
            Top             =   2400
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "Platinas.frx":40AD
            Accelerator     =   83
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameHojalatas 
         Caption         =   "Datos del la Hojalata"
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
         TabIndex        =   0
         Top             =   840
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "DESARROLLO"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   5
            Top             =   2040
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "ESPESOR"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   6
            Top             =   2400
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "BARNIZ"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   4
            ToolTipText     =   "Doble click o signo '+' para ayuda"
            Top             =   1620
            Width           =   480
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "ALTURA"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   3
            Top             =   1245
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "DESCRIP"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   2
            Top             =   840
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "PLATINA"
            DataSource      =   "DataHojalatas"
            Height          =   285
            Index           =   0
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   1
            Top             =   480
            Width           =   1980
         End
         Begin VB.Label LblBarniz 
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
            Left            =   2040
            TabIndex        =   7
            Top             =   1680
            Width           =   5655
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Desarrolllo"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   30
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Espesor"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   29
            Top             =   2400
            Width           =   570
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Barniz"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   28
            Top             =   1665
            Width           =   435
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Altura"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   27
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   26
            Top             =   900
            Width           =   840
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
            TabIndex        =   25
            Top             =   525
            Width           =   600
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "Platinas.frx":43C7
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
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
      Caption         =   "Editar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "Platinas.frx":4909
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
      TabIndex        =   10
      Top             =   4920
      Width           =   1335
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "Platinas.frx":4E4B
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
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "2566;1085"
      Picture         =   "Platinas.frx":538D
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
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "Platinas.frx":58CF
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
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "2355;1085"
      Picture         =   "Platinas.frx":5E11
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Platinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim BBarniz As Boolean
Dim RBuscaBarniz As Recordset
Dim RBuscaAmapro As Recordset

Dim VCodigo As String
Dim VDescripcion As String


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataHojalatas.Recordset
    
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
                        TxtTexto.Item(2).Text = 0
                        TxtTexto.Item(3).Text = 0
                        TxtTexto.Item(5).Text = 0
                        
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
                If Not IsNumeric(TxtTexto.Item(2).Text) Then
                        MsgBox "La Altura Debe Ser Numerico", vbOKOnly + vbInformation, "informacion"
                        TxtTexto.Item(2).Text = "0"
                        TxtTexto.Item(2).SetFocus
                        Exit Sub
                End If
                
                If Not IsNumeric(TxtTexto.Item(3).Text) Then
                        MsgBox "El Desarrollo Debe Ser Numerico", vbOKOnly + vbInformation, "informacion"
                        TxtTexto.Item(3).Text = "0"
                        TxtTexto.Item(3).SetFocus
                        Exit Sub
                End If
                
                        VCodigo = TxtTexto.Item(0).Text
                        VDescripcion = TxtTexto.Item(1).Text
                                
                .Update
                
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                        
                'BUSCA EL CODIGO EN AMAPRO Y SI LO ENCUENTRA LO MODIFICA SINO LO AGREGA
                    Set RBuscaAmapro = Db.OpenRecordset("Select * From CorrelativosMateriaPrima Where Codigomateriaprima = '" & VCodigo & "'")
                        If RBuscaAmapro.RecordCount > 0 Then
                            RBuscaAmapro.Edit
                                RBuscaAmapro!CodigoMateriaPrima = VCodigo
                                RBuscaAmapro!Descripcion = VDescripcion
                            RBuscaAmapro.Update
                        Else
                                RBuscaAmapro.AddNew
                                    RBuscaAmapro!CodigoMateriaPrima = VCodigo
                                    RBuscaAmapro!Descripcion = VDescripcion
                                    RBuscaAmapro!Correlativo = 0
                                    RBuscaAmapro!Espesor = 0
                                    RBuscaAmapro!Minimo = 0
                                RBuscaAmapro.Update
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
                        DataHojalatas.RecordSource = ("Select * From Platinas Where Platina like '" & TxtBusqueda.Text & "*'")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataHojalatas.RecordSource = ("Select * From Platinas where Descrip like '" & TxtBusqueda.Text & "*'")
                    End If
                    DataHojalatas.Refresh
                    DGridHojalatas.Refresh
                    TabHojalatas.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataHojalatas.RecordSource = "Select * From Platinas"
                    DataHojalatas.Refresh
                    DGridHojalatas.Refresh
                    TabHojalatas.Tab = 1
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
         FrameHojalatas.Enabled = True
         DataHojalatas.Visible = False
         DGridHojalatas.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameHojalatas.Enabled = False
         DataHojalatas.Visible = True
         DGridHojalatas.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
    
        TxtTexto.Item(3).Text = DBGridBusqueda.Columns(0).Text
        TxtTexto.Item(3).SetFocus
        FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
            TxtTexto.Item(3).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(3).SetFocus
            FrameBusqueda.Visible = False
    End If
End Sub

Private Sub dgridhojalatas_HeadClick(ByVal ColIndex As Integer)
        DataHojalatas.RecordSource = "Select * From Platinas Order by " & DGridHojalatas.Columns(ColIndex).DataField
        DataHojalatas.Refresh
        DGridHojalatas.Refresh
End Sub

Private Sub Form_Load()
    DataHojalatas.Connect = GConnect
    DataBusqueda.Connect = GConnect
    
    DataHojalatas.DatabaseName = BasedeDatos
    DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Codigo"
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Descripcion"
    End If
    
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
        If Index = 3 Then
            Set RBuscaBarniz = Db.OpenRecordset("Select Descrip From Barniz Where Barniz = '" & TxtTexto.Item(3).Text & "'")
                If RBuscaBarniz.RecordCount > 0 Then
                    LblBarniz.Caption = RBuscaBarniz!Descrip
                Else
                    LblBarniz.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    If Index = 3 Then
        DataBusqueda.RecordSource = "Select * From Barniz"
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
            If Index = 3 Then
                DataBusqueda.RecordSource = "Select * From Barniz"
            End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                DBGridBusqueda.SetFocus
        End If
End Sub
