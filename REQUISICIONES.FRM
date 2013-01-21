VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Requisiciones 
   Caption         =   "Requisiciones De Materia Prima"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "Requisiciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrReportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Frame FrameBuscar 
      Caption         =   "Buscar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      Begin VB.Frame FrameTipos 
         Caption         =   "Tipos De Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5760
         TabIndex        =   40
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
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
            Height          =   195
            Left            =   2280
            TabIndex        =   23
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
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
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10440
         Picture         =   "Requisiciones.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Productos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "InventarioMateriaPrima"
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "Requisiciones.frx":0884
         Height          =   5292
         Left            =   120
         OleObjectBlob   =   "Requisiciones.frx":089D
         TabIndex        =   19
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11052
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado de Requisicion"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8280
         Picture         =   "Requisiciones.frx":1275
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   480
         Left            =   1680
         Picture         =   "Requisiciones.frx":17A7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   480
         Left            =   6960
         Picture         =   "Requisiciones.frx":1CD9
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   9600
         Picture         =   "Requisiciones.frx":220B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salida"
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   480
         Left            =   5640
         Picture         =   "Requisiciones.frx":264D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4320
         Picture         =   "Requisiciones.frx":2B7F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3000
         Picture         =   "Requisiciones.frx":30B1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   480
         Left            =   360
         Picture         =   "Requisiciones.frx":35E3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   1300
      End
      Begin VB.Frame FrameCompras 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   10935
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataRequisiciones"
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBodega 
            Appearance      =   0  'Flat
            DataField       =   "BodegaSalida"
            DataSource      =   "DataRequisiciones"
            Height          =   285
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            DataField       =   "Usuario"
            DataSource      =   "DataRequisiciones"
            Height          =   285
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataRequisiciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   8760
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label LblBodega 
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
            Left            =   3000
            TabIndex        =   41
            Top             =   720
            Width           =   7815
         End
         Begin VB.Label Label6 
            Caption         =   "Bodega Salida"
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
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "No. Requisicion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6480
            TabIndex        =   38
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha"
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
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle de Requisicion"
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
      ForeColor       =   &H00FF0000&
      Height          =   5775
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   11325
      Begin MSDBGrid.DBGrid DBGridDetalleRequisiciones 
         Bindings        =   "Requisiciones.frx":3B15
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "Requisiciones.frx":3B3C
         TabIndex        =   42
         Top             =   1320
         Width           =   11055
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   495
         Left            =   4440
         Picture         =   "Requisiciones.frx":488A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5160
         Width           =   2000
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6600
         Picture         =   "Requisiciones.frx":4DBC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5160
         Width           =   2000
      End
      Begin VB.CommandButton CmdTerminar 
         Caption         =   "&Terminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         Picture         =   "Requisiciones.frx":52EE
         TabIndex        =   17
         Top             =   5160
         Width           =   2000
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2280
         Picture         =   "Requisiciones.frx":5820
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5160
         Width           =   2000
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   495
         Left            =   120
         Picture         =   "Requisiciones.frx":5D52
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5160
         Width           =   2000
      End
      Begin VB.Frame FrameDetalleCompras 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   11055
         Begin MSMask.MaskEdBox TxtCanMatPri 
            DataField       =   "Cantidad"
            DataSource      =   "DataDetalleRequisiciones"
            Height          =   255
            Left            =   9600
            TabIndex        =   12
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBod 
            Appearance      =   0  'Flat
            DataField       =   "Bodega"
            DataSource      =   "DataDetalleRequisiciones"
            Height          =   285
            Left            =   6000
            MaxLength       =   3
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataDetalleRequisiciones"
            Height          =   285
            Left            =   7200
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtCodMatPri 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataDetalleRequisiciones"
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox TxtDesPro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   480
            Width           =   7575
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad"
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
            Left            =   9600
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Data DataDetalleRequisiciones 
      Caption         =   "Detalle Requisiciones"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleRequisicion"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataRequisiciones 
      Caption         =   "Requisiciones"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoRequisicion"
      Top             =   8145
      Width           =   11190
   End
End
Attribute VB_Name = "Requisiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String
Dim VDocumento As Double
Dim VDocumentoDetalle As Double
Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim BBodega As Boolean
Dim BMateriaPrima As Boolean

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaSigDoc As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaBodega As Recordset




Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         DataRequisiciones.Visible = False
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         DataRequisiciones.Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If

End Sub



Private Sub CmdAgregar2_Click()
On Error Resume Next
    'AGREGA UN REGISTRO DE DETALLE
    DataDetalleRequisiciones.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description
    End If
        
    Bandera2 = True
    Botones2

    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleRequisiciones.Enabled = False
    'SE LE ASIGNA A LA BODEGA DETALLE LA BODEGA DEL ENCABEZADO
    TxtBod.Text = VBodega
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    TxtDocDet.Text = VDocumento
    TxtCodMatPri.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

VDocumento = TxtDoc.Text

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If Mensaje = vbOK Then
                MousePointer = 11
                
                'SE BUSCAN TODO EL DETALLE DE LA REQUISICION PARA PODER SUMAR DE NUEVO LA EXISTENCIA
                Set RBuscaDetalle = Db.OpenRecordset("Select Codigo, Cantidad, Bodega from DetalleRequisicion where Documento = " & VDocumento)
                If RBuscaDetalle.RecordCount > 0 Then
                    Do Until RBuscaDetalle.EOF
                        'BUSCAMOS LA MATERIA PRIMA PARA REGRESAR LAS CANTIDADES Y SUMAR EXISTENCIAS
                        Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & RBuscaDetalle!Codigo & "' and Bodega = '" & RBuscaDetalle!Bodega & "'")
                            RBuscaMateriaPrima.Edit
                                RBuscaMateriaPrima!Existencia = RBuscaMateriaPrima!Existencia + RBuscaDetalle!Cantidad
                                RBuscaMateriaPrima!Salidas = RBuscaMateriaPrima!Salidas - RBuscaDetalle!Cantidad
                            RBuscaMateriaPrima.Update
                        RBuscaDetalle.MoveNext
                    Loop
                End If
                
                'BORRA EL ENCABEZADO DE LA FACTURA
                DataRequisiciones.Recordset.Delete
                DataRequisiciones.Recordset.MoveLast
                MousePointer = 0
            End If
  
            If DataRequisiciones.Recordset.EOF Then
                DataRequisiciones.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next

    'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
    VDocumentoDetalle = TxtDocDet.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
    VCantidadMateriaPrima = TxtCanMatPri.Text

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If Mensaje = vbOK Then
                MousePointer = 11
                
                    'BUSCAMOS LA MATERIA PRIMA SUMAMOS LA EXISTENCIA Y RESTAMOS LAS SALIDAS
                    Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and bodega = '" & VBodega & "'")
                        RBuscaMateriaPrima.Edit
                            RBuscaMateriaPrima!Existencia = RBuscaMateriaPrima!Existencia + VCantidadMateriaPrima
                            RBuscaMateriaPrima!Salidas = RBuscaMateriaPrima!Salidas - VCantidadMateriaPrima
                        RBuscaMateriaPrima.Update
                        
                        If Err <> 0 Then
                        End If
                        
                   'BORRA EL DETALLE DE LA FACTURA
                    DataDetalleRequisiciones.Recordset.Delete
                    
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                        End If
                                        
                    'SELECCIONA TODOS LOS DETALLES DE LA FACTURA
                    DataDetalleRequisiciones.RecordSource = ("Select * from DetalleRequisicion where documento = " & VDocumentoDetalle & " order By Codigo")
                    DataDetalleRequisiciones.Refresh
                    DBGridDetalleRequisiciones.Refresh
                MousePointer = 0
            End If
  
            If DataRequisiciones.Recordset.EOF Then
                DataRequisiciones.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBuscar_Click()
    Mensaje = InputBox("No. Requisicion a Buscar")
    DataRequisiciones.Recordset.FindFirst ("Documento = " & Mensaje)
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next

    'GRABA DATOS
    DataRequisiciones.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA EL INGRESO
    DataDetalleRequisiciones.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    DBGridDetalleRequisiciones.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    'EDITA EL REGISTRO ACTUAL
    DataRequisiciones.Recordset.Edit
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
End Sub


Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    'GUARDA VARIABLES
    VCantidadMateriaPrima = TxtCanMatPri.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
            
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanMatPri.Text) Then
       MsgBox "Cantidad De Materia Prima Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanMatPri.SetFocus
       Exit Sub
    End If
    
    'REVISA SI EXISTE LA MATERIA PRIMA
    Set RBuscaMateriaPrima = Db.OpenRecordset("Select * from InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
    If RBuscaMateriaPrima.RecordCount > 0 Then
    Else
        MsgBox "Codigo De Materia Prima No Existe En La Bodega Que Esta Trabajando", vbOKOnly + vbCritical, "Error"
        TxtCodMatPri.SetFocus
        Exit Sub
    End If
        
    'GRABA DATOS
    DataDetalleRequisiciones.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera2 = False
    Botones2
    
     'BUSCA LA EXISTENCIA Y LAS SALIDAS DE LA MATERIA PRIMA Y RESTA LA EXISTENCIA Y SUMA SALIDAS
    Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
        If RBuscaMateriaPrima.RecordCount > 0 Then
                RBuscaMateriaPrima.Edit
                    RBuscaMateriaPrima!Existencia = (RBuscaMateriaPrima!Existencia - VCantidadMateriaPrima)
                    RBuscaMateriaPrima!Salidas = (RBuscaMateriaPrima!Salidas + VCantidadMateriaPrima)
                RBuscaMateriaPrima.Update
        End If
  
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleRequisiciones.RecordSource = ("Select * from DetalleRequisicion where Documento = " & VDocumento & " Order BY Codigo")
    DataDetalleRequisiciones.Refresh
    DBGridDetalleRequisiciones.Refresh
        
    DBGridDetalleRequisiciones.Enabled = True
    CmdAgregar2_Click
End Sub


Private Sub CmdAgregar_Click()
    On Error Resume Next
    'AGREGA UN REGISTRO
    DataRequisiciones.Recordset.AddNew
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    'ASIGNA EL USUARIO
    TxtUsu.Text = GUsuario
        
    MskFec.Text = Date
    MskFec.SetFocus
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoRequisicion")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                TxtDoc.Text = "1"
            Else
                TxtDoc.Text = RBuscaSigDoc(0) + 1
            End If
        End If
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    VDocumento = TxtDoc.Text
    VBodega = TxtBodega.Text
    
    If Not IsNumeric(TxtDoc.Text) Then
        MsgBox "Numero De Requisicion Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
    
    
    'GRABA DATOS
    DataRequisiciones.Recordset.Update
    
    'SI SE DUPLICA LA LLAVE
    If Err = 3022 Then
        MsgBox "Numero De Requisicion Ya Existe", vbOKOnly + vbCritical, "Error"
        TxtDoc.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    'SELECCIONA TODO EL DETALLE DE LA REQUISICION
    DataDetalleRequisiciones.RecordSource = ("Select * from DetalleRequisicion where Documento = " & VDocumento & " Order by Codigo")
    DataDetalleRequisiciones.Refresh
    DBGridDetalleRequisiciones.Refresh
    
    'MUEVE EL RECORDSET A LA FACTURA ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataRequisiciones.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = False
    CmdAgregar2_Click
End Sub

Private Sub CmdImprimir_Click()
'MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDoc.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                'CrReportes.SelectionFormula = "UpperCase({EncabezadoSalidas.Documento}) = '" & TxtDoc.Text & "'"
                'CrReportes.ReportFileName = App.Path & "\Salidas.rpt"
                'CrReportes.Action = 1
'MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True

End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub DBGridBuscar_DblClick()
        If BMateriaPrima = True Then
            TxtCodMatPri.Text = DBGridBuscar.Columns(0)
        ElseIf BBodega = True Then
            TxtBodega.Text = DBGridBuscar.Columns(0)
            TxtBodega.SetFocus
        End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtCodMatPri.Text = DBGridBuscar.Columns(0)
            ElseIf BBodega = True Then
                TxtBodega.Text = DBGridBuscar.Columns(0)
                TxtBodega.SetFocus
            End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
        End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
    'ASIGNAMOS A QUE BASE DE DATOS SE VAN A CONECTAR LOS DATAS
    DataRequisiciones.DatabaseName = BasedeDatos
    DataDetalleRequisiciones.DatabaseName = BasedeDatos
    DataBuscar.DatabaseName = BasedeDatos
    
    
    
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


Private Sub TxtBodega_Change()
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodega.Text & "'")
        If RBuscaBodega.RecordCount > 0 Then
            LblBodega.Caption = RBuscaBodega!Descripcion
        Else
            LblBodega.Caption = ""
        End If
    
End Sub

Private Sub TxtBodega_DblClick()
        BMateriaPrima = False
        BBodega = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA LAS BODEGAS DE MATERIA PRIMA
        DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub

Private Sub TxtBodega_GotFocus()
        TxtBodega.SelStart = 0
        TxtBodega.SelLength = Len(TxtBodega.Text)
End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BMateriaPrima = False
        BBodega = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA LAS BODEGAS DE MATERIA PRIMA
        DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If

End Sub

Private Sub TxtBuscar_Change()
If BMateriaPrima = True Then
    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
    If OptCodigo.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And CodigoMateriaPrima Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        Else
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        End If
    ElseIf OptDescripcion.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
                
        Else
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        End If
    End If
ElseIf BBodega = True Then
    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
    If OptCodigo.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '" & TxtBuscar.Text & "*'")
        Else
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '*" & TxtBuscar.Text & "*'")
        End If
    ElseIf OptDescripcion.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'")
                
        Else
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'")
        End If
    End If
End If
    DataBuscar.Refresh
    DBGridBuscar.Refresh
    DBGridBuscar.Columns(1).Width = "4000"

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

Private Sub txtCanMatPri_GotFocus()
        TxtCanMatPri.SelStart = 0
        TxtCanMatPri.SelLength = Len(TxtCanMatPri.Text)
End Sub
Private Sub txtCanMatPri_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub
Private Sub TxtCodMatPri_Change()
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From InventarioMateriaPrima Where CodigoMateriaPrima = '" & TxtCodMatPri.Text & "' And Bodega = '" & TxtBod.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                Else
                        TxtDesPro.Text = ""
                End If
End Sub
Private Sub TxtCodMatPri_DblClick()
        BMateriaPrima = True
        BBodega = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' Order by CodigoMateriaPrima")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub
Private Sub TxtCodMatPri_GotFocus()
        TxtCodMatPri.SelStart = 0
        TxtCodMatPri.SelLength = Len(TxtCodMatPri.Text)
End Sub

Private Sub TxtCodMatPri_KeyPress(KeyAscii As Integer)
        'SI PRECIONA ENTER
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        'SI PRECIONA LA TECLA DE SIGNO +
        If KeyAscii = 43 Then
            BMateriaPrima = True
            BBodega = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' Order by CodigoMateriaPrima")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
        End If
End Sub


Private Sub TxtDesPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

