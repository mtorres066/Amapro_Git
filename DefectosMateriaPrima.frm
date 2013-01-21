VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DefectosMateriaPrima 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Defectos Materia Prima"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "DefectosMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
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
      Height          =   6135
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
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
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "DefectosMateriaPrima.frx":6296
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "DefectosMateriaPrima.frx":8308
         Height          =   4935
         Left            =   120
         OleObjectBlob   =   "DefectosMateriaPrima.frx":8323
         TabIndex        =   28
         ToolTipText     =   "Signo '+' O Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   8175
      End
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   4695
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "DefectosMateriaPrima.frx":8CFD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDefectos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "DefectosMateriaPrima.frx":9017
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridDefectos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "DefectosMateriaPrima.frx":9469
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBuscar(1)"
      Tab(2).Control(1)=   "CmdBuscar(0)"
      Tab(2).Control(2)=   "Txtbuscar"
      Tab(2).Control(3)=   "OptNumero"
      Tab(2).Control(4)=   "OptCodigo"
      Tab(2).Control(5)=   "Lblbusqueda2"
      Tab(2).Control(6)=   "LblBusqueda"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   735
         Index           =   1
         Left            =   -68760
         Picture         =   "DefectosMateriaPrima.frx":98BB
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   735
         Index           =   0
         Left            =   -68760
         Picture         =   "DefectosMateriaPrima.frx":9BC5
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.OptionButton OptNumero 
         Caption         =   "Numero Ingreso"
         Height          =   195
         Left            =   -72840
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Materia Prima"
         Height          =   195
         Left            =   -74640
         TabIndex        =   17
         Top             =   1200
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Frame FrameDefectos 
         Caption         =   "Defectos Del Bulto"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   8115
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Defecto"
            DataSource      =   "DataDefectos"
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataDefectos"
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoMateriaPrima"
            DataSource      =   "DataDefectos"
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label LblDefecto 
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
            TabIndex        =   16
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Label LblMateriaPrima 
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
            TabIndex        =   15
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label2 
            Caption         =   "Defecto"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label2 
            Caption         =   "# Ingreso"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGridDefectos 
         Bindings        =   "DefectosMateriaPrima.frx":A007
         Height          =   3825
         Left            =   -74880
         OleObjectBlob   =   "DefectosMateriaPrima.frx":A022
         TabIndex        =   9
         Top             =   720
         Width           =   8145
      End
      Begin VB.Label Lblbusqueda2 
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
         Left            =   -71040
         TabIndex        =   21
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label LblBusqueda 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Materia Prima"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   2520
         Width           =   1815
      End
   End
   Begin VB.Data DataDefectos 
      BackColor       =   &H80000014&
      Caption         =   "Defectos De Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DefectosMateriaPrima"
      Top             =   5760
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "DefectosMateriaPrima.frx":ABD8
      Picture         =   "DefectosMateriaPrima.frx":B01A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "DefectosMateriaPrima.frx":B45C
      Picture         =   "DefectosMateriaPrima.frx":B89E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "DefectosMateriaPrima.frx":BDD0
      Picture         =   "DefectosMateriaPrima.frx":C212
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "DefectosMateriaPrima.frx":C744
      Picture         =   "DefectosMateriaPrima.frx":CB86
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "DefectosMateriaPrima.frx":D0B8
      Picture         =   "DefectosMateriaPrima.frx":D4FA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "DefectosMateriaPrima.frx":DA2C
      Picture         =   "DefectosMateriaPrima.frx":DE6E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1200
   End
End
Attribute VB_Name = "DefectosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim BMateriaPrima As Boolean
Dim BDefecto As Boolean

Dim mensaje As String
Dim buscar As String

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaDefectos As Recordset

Dim RBuscaBulto As Recordset

Sub botones()
    If Bandera = True Then
         FrameDefectos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         DataDefectos.Visible = False
         DbGridDefectos.Visible = False
    Else
         FrameDefectos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataDefectos.Visible = True
         DbGridDefectos.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataDefectos.Recordset
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
                    Txttexto.Item(0).SetFocus
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
                    If Not IsNumeric(Txttexto.Item(1).Text) Then
                        MsgBox "Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(1).SetFocus
                        Exit Sub
                    End If
                    
                    'BUSCA SI EXISTE EL BULTO
                    Set RBuscaBulto = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Codigo = '" & Txttexto.Item(0).Text & "' And NumeroIngreso = " & Txttexto.Item(1).Text)
                        If RBuscaBulto.RecordCount > 0 Then
                        Else
                            MsgBox "Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            Txttexto.Item(1).SetFocus
                            Exit Sub
                        End If
                    
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                      'SI ES CUALQUIER OTRO ERROR
                     If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                     End If
                        Bandera = False
                        botones
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
                        DataDefectos.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataDefectos.Recordset.MoveLast
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataDefectos.Recordset.EOF Then
                        DataDefectos.Recordset.MoveLast
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
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                DataDefectos.RecordSource = "Select * From DefectosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*'"
            ElseIf OptNumero.Value = True Then
                If IsNumeric(TxtBusqueda.Text) Then
                    DataDefectos.RecordSource = "Select * From DefectosMateriaPrima Where NumeroIngreso = " & TxtBuscar.Text
                Else
                    MsgBox "Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                    TxtBusqueda.SetFocus
                    Exit Sub
                End If
            End If
        'SELECCIONAR TODOS
        ElseIf Index = 1 Then
                DataDefectos.RecordSource = "Select * From DefectosMateriaPrima"
        End If
                DataDefectos.Refresh
                DbGridDefectos.Refresh
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BMateriaPrima = True Then
            Txttexto.Item(0).Text = DBGridBusqueda.Columns(0)
            Txttexto.Item(0).SetFocus
        ElseIf BDefecto = True Then
            Txttexto.Item(2).Text = DBGridBusqueda.Columns(0)
            Txttexto.Item(2).SetFocus
        End If
            FrameBusqueda.Visible = False
            
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                Txttexto.Item(0).Text = DBGridBusqueda.Columns(0)
                Txttexto.Item(0).SetFocus
            ElseIf BDefecto = True Then
                Txttexto.Item(2).Text = DBGridBusqueda.Columns(0)
                Txttexto.Item(2).SetFocus
            End If
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub Dbgriddefectos_HeadClick(ByVal ColIndex As Integer)
    DataDefectos.RecordSource = ("Select * from DefectosMateriaPrima order by " & DbGridDefectos.Columns(ColIndex).DataField)
    DataDefectos.Refresh
    DbGridDefectos.Refresh
End Sub

Private Sub Form_Load()
        DataDefectos.ConnectionString = GTipoProveedor
        DataBusqueda.ConnectionString = GTipoProveedor
        
        DataDefectos.Refresh
        DataBusqueda.Refresh
End Sub


Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub Txtbusqueda_Change()
            
    'MATERIA PRIMA
    If BMateriaPrima = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima where Descripcion Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima where Descripcion Like '*" & TxtBusqueda.Text & "*'"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima where CodigoMateriaPrima Like '" & TxtBusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima where CodigoMateriaPrima Like '*" & TxtBusqueda.Text & "*'"
                End If
            End If
    'DEFECTOS
    ElseIf BDefecto = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos where Descrip Like '" & TxtBusqueda.Text & "*' Order By Descrip"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos where Descrip Like '*" & TxtBusqueda.Text & "*' Order By Descrip"
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptTipo.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos where Defecto Like '" & TxtBusqueda.Text & "*' Order By Descrip"
                'CUALQUIER PALABRA
                ElseIf OptTipo.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos where Defecto Like '*" & TxtBusqueda.Text & "*' Order By Descrip"
                End If
            End If
    End If
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBuscar_GotFocus()
        TxtBuscar.SelStart = 0
        TxtBuscar.SelLength = Len(TxtBuscar.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 0 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & Txttexto.Item(0).Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblMateriaPrima.Caption = ""
                End If
        ElseIf Index = 2 Then
            Set RBuscaDefectos = Db.OpenRecordset("Select Descrip From Defectos Where Defecto = '" & Txttexto.Item(2).Text & "'")
                If RBuscaDefectos.RecordCount > 0 Then
                    LblDefecto.Caption = RBuscaDefectos!Descrip
                Else
                    LblDefecto.Caption = ""
                End If
        End If
        
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 0 Then
            BMateriaPrima = True
            BDefecto = False
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima"
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
        ElseIf Index = 2 Then
            BMateriaPrima = False
            BDefecto = True
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos"
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
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
            BMateriaPrima = True
            BDefecto = False
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima"
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
        ElseIf Index = 2 Then
            BMateriaPrima = False
            BDefecto = True
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DataBusqueda.RecordSource = "Select Defecto, Descrip From Defectos"
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
        End If
    End If
End Sub
