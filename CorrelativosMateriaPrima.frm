VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CorrelativosMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Catalogo De Articulos"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "CorrelativosMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTiposMP 
      Caption         =   "Tipos De Catalogo"
      Height          =   7575
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Data DataTiposMP 
         Caption         =   "Tipos De Mp"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TiposDeMateriaPrima"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7680
         Picture         =   "CorrelativosMateriaPrima.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Salida"
         Top             =   240
         Width           =   615
      End
      Begin MSDBGrid.DBGrid DBGridTiposMP 
         Bindings        =   "CorrelativosMateriaPrima.frx":237C
         Height          =   7095
         Left            =   120
         OleObjectBlob   =   "CorrelativosMateriaPrima.frx":2396
         TabIndex        =   35
         ToolTipText     =   "Signo '+' o Doble Click para seleccionar"
         Top             =   240
         Width           =   7455
      End
   End
   Begin TabDlg.SSTab TabCorrelativos 
      Height          =   6135
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CorrelativosMateriaPrima.frx":2D8F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCorrelativos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CorrelativosMateriaPrima.frx":30A9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CorrelativosMateriaPrima.frx":34FB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones2"
      Tab(2).Control(1)=   "CmdActualizar"
      Tab(2).Control(2)=   "CmdBusqueda"
      Tab(2).Control(3)=   "TxtBuscar"
      Tab(2).Control(4)=   "FrameOpciones"
      Tab(2).Control(5)=   "Lbletiqueta"
      Tab(2).ControlCount=   6
      Begin VB.Frame FrameOpciones2 
         Caption         =   "Tipo De Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -69600
         TabIndex        =   43
         Top             =   960
         Width           =   2775
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   435
            Left            =   1680
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   435
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Left            =   -69240
         Picture         =   "CorrelativosMateriaPrima.frx":394D
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Seleccionar Datos"
         Height          =   855
         Left            =   -69240
         Picture         =   "CorrelativosMateriaPrima.frx":3C57
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -69240
         TabIndex        =   20
         ToolTipText     =   " "
         Top             =   3720
         Width           =   2445
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   960
         Width           =   3885
         Begin VB.OptionButton OptTipo 
            Caption         =   "Tipo Catalogo"
            Height          =   195
            Left            =   2400
            TabIndex        =   42
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame FrameCorrelativos 
         Caption         =   "Datos de Articulos"
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
         Height          =   5055
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   8115
         Begin VB.TextBox TxtPesUni 
            Appearance      =   0  'Flat
            DataField       =   "PesoxUnidad"
            DataSource      =   "DataCorrelativos"
            Height          =   288
            Left            =   2400
            TabIndex        =   9
            Top             =   3600
            Width           =   1932
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Es Ficha Tecnica ?"
            DataField       =   "EsFichaTecnica"
            DataSource      =   "DataCorrelativos"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   4680
            Width           =   2535
         End
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   4680
            Width           =   1935
         End
         Begin VB.TextBox TxtCueLam 
            Appearance      =   0  'Flat
            DataField       =   "CuerposPorLamina"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            TabIndex        =   8
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox TxtEsp 
            Appearance      =   0  'Flat
            DataField       =   "Espesor"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Top             =   2520
            Width           =   1935
         End
         Begin VB.TextBox TxtMin 
            Appearance      =   0  'Flat
            DataField       =   "Minimo"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            TabIndex        =   7
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox TxtTipMatPri 
            Appearance      =   0  'Flat
            DataField       =   "TipoDeMateriaPrima"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   5
            ToolTipText     =   "Signo '+' o Doble Click para Ayuda"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox TxtUniMedPes 
            Appearance      =   0  'Flat
            DataField       =   "UnidadMedidaPeso"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox TxtUniMed 
            Appearance      =   0  'Flat
            DataField       =   "UnidadMedida"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            DataField       =   "Descripcion"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   5535
         End
         Begin MSMask.MaskEdBox MskCor 
            DataField       =   "Correlativo"
            DataSource      =   "DataCorrelativos"
            Height          =   288
            Left            =   2400
            TabIndex        =   2
            Top             =   1080
            Width           =   1932
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoMateriaPrima"
            DataSource      =   "DataCorrelativos"
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso x Unidad En Kilos"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   46
            Top             =   3600
            Width           =   1650
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   7
            Left            =   5280
            TabIndex        =   41
            Top             =   4680
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuerpos x Lamina"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   3240
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Minimo de Inventario"
            Height          =   192
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   2880
            Width           =   1476
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Espesor"
            Height          =   192
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   2520
            Width           =   576
         End
         Begin VB.Label LblTipoDeMateriaPrima 
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
            Height          =   252
            Left            =   4440
            TabIndex        =   33
            Top             =   2160
            Width           =   3492
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo De Catalogo"
            Height          =   192
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   1248
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidad Medida Peso"
            Height          =   192
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   1488
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidad Medida Bulto"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1488
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo De Articulo"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Numero Consecutivo Maximo"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   2076
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CorrelativosMateriaPrima.frx":4099
         Height          =   5265
         Left            =   -74880
         OleObjectBlob   =   "CorrelativosMateriaPrima.frx":40B8
         TabIndex        =   17
         Top             =   720
         Width           =   8145
      End
      Begin VB.Label Lbletiqueta 
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
         Left            =   -71760
         TabIndex        =   28
         Top             =   3720
         Width           =   2415
      End
   End
   Begin VB.Data DataCorrelativos 
      BackColor       =   &H80000014&
      Caption         =   "Correlativos Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\erick\Amapro Metalenvases\metalenvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CorrelativosMateriaPrima"
      Top             =   7200
      Width           =   8175
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6840
      MouseIcon       =   "CorrelativosMateriaPrima.frx":5B93
      Picture         =   "CorrelativosMateriaPrima.frx":5FD5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1320
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5520
      MouseIcon       =   "CorrelativosMateriaPrima.frx":8047
      Picture         =   "CorrelativosMateriaPrima.frx":8489
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4200
      MouseIcon       =   "CorrelativosMateriaPrima.frx":89BB
      Picture         =   "CorrelativosMateriaPrima.frx":8DFD
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   2880
      MouseIcon       =   "CorrelativosMateriaPrima.frx":932F
      Picture         =   "CorrelativosMateriaPrima.frx":9771
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1560
      MouseIcon       =   "CorrelativosMateriaPrima.frx":9CA3
      Picture         =   "CorrelativosMateriaPrima.frx":A0E5
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   240
      MouseIcon       =   "CorrelativosMateriaPrima.frx":A617
      Picture         =   "CorrelativosMateriaPrima.frx":AA59
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1200
   End
End
Attribute VB_Name = "CorrelativosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean

Dim mensaje As String
Dim buscar As String
Dim VCodigoViejo As String
Dim VCodigoNuevo As String
Dim VDescripcion As String
Dim VPeso As Single
Dim VUnidadesxLamina As Integer

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaFichaTecnica As Recordset


Sub botones()
    If Bandera = True Then
         FrameCorrelativos.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataCorrelativos.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameCorrelativos.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataCorrelativos.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub



Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub CmdActualizar_Click()
    DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima")
    DataCorrelativos.Refresh
    DBGrid1.Refresh
    TabCorrelativos.Tab = 1
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        DataCorrelativos.Recordset.AddNew
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        Bandera = True
        botones
        TxtCod.SetFocus
        TxtUsuario.Text = GUsuario
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataCorrelativos.Recordset.Delete
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                DataCorrelativos.Recordset.MoveLast
            End If
  
            If DataCorrelativos.Recordset.EOF Then
                DataCorrelativos.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdBusqueda_Click()
        'CODIGO
        If OptCodigo.Value = True Then
            'CUALQUIER PALAPRA
            If OptCuaPal.Value = True Then
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where CodigoMateriaPrima like '*" & TxtBuscar.Text & "*'")
            'PALABRA INICIAL
            Else
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where CodigoMateriaPrima like '" & TxtBuscar.Text & "*'")
            End If
        'DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            'CUALQUIER PALAPRA
            If OptCuaPal.Value = True Then
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where Descripcion like '*" & TxtBuscar.Text & "*'")
            'PALABRA INICIAL
            Else
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where Descripcion like '" & TxtBuscar.Text & "*'")
            End If
        'TIPO DE MATERIA PRIMA
        ElseIf OptTipo.Value = True Then
            'CUALQUIER PALAPRA
            If OptCuaPal.Value = True Then
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where TipoDeMateriaPrima = '" & TxtBuscar.Text & "'")
            'PALABRA INICIAL
            Else
                DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima where TipoDeMateriaPrima = '" & TxtBuscar.Text & "'")
            End If
        End If
            DataCorrelativos.Refresh
            DBGrid1.Refresh
            TabCorrelativos.Tab = 1
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        DataCorrelativos.Recordset.CancelUpdate
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
        End If
        Bandera = False
        botones
        
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        
        VCodigoViejo = TxtCod
        
        DataCorrelativos.Recordset.Edit
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        Bandera = True
        botones
        TxtCod.SetFocus
        TxtUsuario.Text = GUsuario
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
       
   'REVISA LA DESCRIPCION
   If TxtDes.Text = "" Then
        MsgBox "La Descripcion No Puede Estar en Blanco", vbOKOnly + vbInformation, "Informacion"
        TxtDes.SetFocus
        Exit Sub
   End If
   
   'REVISA CORRELATIVO
   If Not IsNumeric(MskCor.Text) Then
        MsgBox "Numero De Correlativo Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        MskCor.SetFocus
        Exit Sub
   End If
       
   'REVISA ESPESOR
   If Not IsNumeric(TxtEsp.Text) Then
        MsgBox "Espesor Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtEsp.SetFocus
        Exit Sub
   End If
    
   'REVISA MINIMO
   If Not IsNumeric(TxtMin.Text) Then
        MsgBox "Minimo Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtMin.SetFocus
        Exit Sub
   End If
    
   'REVISA CUERPOS POR LAMINA
   If Not IsNumeric(TxtCueLam.Text) Then
        MsgBox "Cuerpos Por Lamina Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtCueLam.SetFocus
        Exit Sub
   End If
    
   VCodigoViejo = TxtCod.Text
   VDescripcion = TxtDes.Text
   VPeso = TxtPesUni.Text
   VUnidadesxLamina = TxtCueLam.Text
   
   DataCorrelativos.Recordset.Update
   
   If Err = 3022 Then
        MsgBox "Codigo de Materia Prima Ya Existe", vbOKOnly + vbInformation, "Informacion"
        TxtCod.SetFocus
   ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
   Else
   
      'BUSCA EL CODIGO EN LAS FICHA TECNICAS
      Set RBuscaFichaTecnica = Db.OpenRecordset("Select Esp_Tec, Descrip, PesoxUnidad, UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & VCodigoViejo & "'")
        If RBuscaFichaTecnica.RecordCount > 0 Then
                RBuscaFichaTecnica.Edit
                        'RBuscaFichaTecnica!Esp_Tec = VCodigoNuevo
                        RBuscaFichaTecnica!Descrip = VDescripcion
                        RBuscaFichaTecnica!PesoxUnidad = VPeso
                        RBuscaFichaTecnica!UnidadesxLamina = VUnidadesxLamina
                RBuscaFichaTecnica.Update
        End If
        
        Bandera = False
        botones
        CmdAgregar.SetFocus
  End If
      
End Sub

Private Sub CmdSale_Click()
    FrameTiposMP.Visible = False
    TxtTipMatPri.SetFocus
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataCorrelativos.RecordSource = ("Select * from CorrelativosMateriaPrima order by " & DBGrid1.Columns(ColIndex).DataField)
    DataCorrelativos.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub DBGridTiposMP_DblClick()
        TxtTipMatPri.Text = DBGridTiposMP.Columns(0)
        FrameTiposMP.Visible = False
        TxtTipMatPri.SetFocus
End Sub

Private Sub DBGridTiposMP_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            TxtTipMatPri.Text = DBGridTiposMP.Columns(0)
            FrameTiposMP.Visible = False
            TxtTipMatPri.SetFocus
        End If
End Sub

Private Sub Form_Load()
    DataCorrelativos.ConnectionString = GTipoProveedor
    DataTiposMP.ConnectionString = GTipoProveedor
    
    DataCorrelativos.Refresh
    DataTiposMP.Refresh
    
    'VALIDA SI EL USUARIO PUEDE EDITAR
    If GEditar = True Then
        DBGrid1.AllowUpdate = True
    Else
        DBGrid1.AllowUpdate = False
    End If
    
End Sub

Private Sub MskCor_GotFocus()
        MskCor.SelStart = 0
        MskCor.SelLength = Len(MskCor.Text)
End Sub

Private Sub MskCor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub OptCodigo_Click()
    Lbletiqueta.Caption = "Codigo"
    TxtBuscar.SetFocus
End Sub

Private Sub OptDescripcion_Click()
    Lbletiqueta.Caption = "Descripcion"
    TxtBuscar.SetFocus
End Sub



Private Sub OptTipo_Click()
    Lbletiqueta.Caption = "Tipo De Catalogo"
    TxtBuscar.SetFocus
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


Private Sub TxtCod_GotFocus()
    TxtCod.SelStart = 0
    TxtCod.SelLength = Len(TxtCod.Text)
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtCueLam_GotFocus()
    TxtCueLam.SelStart = 0
    TxtCueLam.SelLength = Len(TxtCueLam.Text)
End Sub

Private Sub TxtCueLam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtDes_GotFocus()
    TxtDes.SelStart = 0
    TxtDes.SelLength = Len(TxtDes.Text)
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub TxtEsp_GotFocus()
        TxtEsp.SelStart = 0
        TxtEsp.SelLength = Len(TxtEsp.Text)
End Sub

Private Sub TxtEsp_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtMin_GotFocus()
    TxtMin.SelStart = 0
    TxtMin.SelLength = Len(TxtMin.Text)
End Sub

Private Sub TxtMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtTipMatPri_Change()
    Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima where CodigoTipo = '" & TxtTipMatPri.Text & "'")
        If RBuscaMateriaPrima.RecordCount > 0 Then
            LblTipoDeMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
        Else
            LblTipoDeMateriaPrima.Caption = ""
        End If
End Sub

Private Sub TxtTipMatPri_DblClick()
    DataTiposMP.RecordSource = "Select * From TiposDeMateriaPrima"
    DataTiposMP.Refresh
    DBGridTiposMP.Refresh
    FrameTiposMP.Visible = True
    DBGridTiposMP.SetFocus
End Sub

Private Sub TxtTipMatPri_GotFocus()
    TxtTipMatPri.SelStart = 0
    TxtTipMatPri.SelLength = Len(TxtTipMatPri.Text)
End Sub

Private Sub TxtTipMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If KeyAscii = 43 Then
        DataTiposMP.RecordSource = "Select * From TiposDeMateriaPrima"
        DataTiposMP.Refresh
        DBGridTiposMP.Refresh
        FrameTiposMP.Visible = True
        DBGridTiposMP.SetFocus
    End If
End Sub

Private Sub TxtUniMed_GotFocus()
    TxtUniMed.SelStart = 0
    TxtUniMed.SelLength = Len(TxtUniMed.Text)
End Sub

Private Sub TxtUniMed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub TxtUniMedPes_GotFocus()
    TxtUniMedPes.SelStart = 0
    TxtUniMedPes.SelLength = Len(TxtUniMed.Text)
End Sub

Private Sub TxtUniMedPes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
