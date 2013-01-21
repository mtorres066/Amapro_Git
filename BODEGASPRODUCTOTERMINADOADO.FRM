VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BodegasProductoTerminadoADO 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Tecnica De Bodegas Producto Terminado"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "BODEGASPRODUCTOTERMINADOADO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   5280
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"BODEGASPRODUCTOTERMINADOADO.frx":0442
      OLEDBString     =   $"BODEGASPRODUCTOTERMINADOADO.frx":04D0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "administrador"
      Password        =   "ahamay"
      RecordSource    =   "Select * From BodegasProductoTerminado"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   4095
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "BODEGASPRODUCTOTERMINADOADO.frx":055E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "BODEGASPRODUCTOTERMINADOADO.frx":0878
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "BODEGASPRODUCTOTERMINADOADO.frx":0CCA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar(0)"
      Tab(2).Control(4)=   "CmdBuscar(1)"
      Tab(2).ControlCount=   5
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "BODEGASPRODUCTOTERMINADOADO.frx":111C
         Height          =   3255
         Left            =   -74880
         TabIndex        =   25
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "CodigoBodega"
            Caption         =   "CodigoBodega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Direccion"
            Caption         =   "Direccion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Telefono"
            Caption         =   "Telefono"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Encargado"
            Caption         =   "Encargado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":1131
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":143B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -71040
         TabIndex        =   13
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2160
         Width           =   1845
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   960
         Width           =   4005
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   750
            TabIndex        =   11
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   1220
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   2160
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos de Bodega Producto Terminado"
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
         Height          =   2775
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   8115
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "Encargado"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   4
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "Telefono"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   3
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            DataField       =   "Direccion"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   6855
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoBodega"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   6855
         End
         Begin VB.Label Label2 
            Caption         =   "Encargado"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Telefono"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Direccion"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
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
         Left            =   -73080
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":187D
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":1CBF
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":2101
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":2543
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":2A75
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":2EB7
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":33E9
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":382B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":3D5D
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":419F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "BODEGASPRODUCTOTERMINADOADO.frx":46D1
      Picture         =   "BODEGASPRODUCTOTERMINADOADO.frx":4B13
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " "
      Top             =   4320
      Width           =   1200
   End
End
Attribute VB_Name = "BodegasProductoTerminadoADO"
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
         FrameBodegas.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Txttexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         Adodc1.Visible = False
         FrameOpciones.Visible = False
         DataGrid1.Visible = False
    Else
         FrameBodegas.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         Adodc1.Visible = True
         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'On Error Resume Next
    '    Adodc1.Caption = "Registro Actual " & pRecordset.AbsolutePosition
   '
   '     If Err Or pRecordset.BOF Or pRecordset.EOF Then
   '         Adodc1.Caption = "Ningun Registro Activo"
   '         Adodc1.Recordset.MoveLast
   '     End If
   '
   '     Err = 0
            
End Sub

Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With Adodc1.Recordset
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
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "Codigo de Bodega ya existe", vbOKOnly + vbInformation, "Informacion"
                        Txttexto.Item(0).SetFocus
                      'SI ES CUALQUIER OTRO ERROR
                     ElseIf Err <> 3022 And Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                     End If
                        Bandera = False
                        botones
            'CANCELAR
            ElseIf Index = 3 Then
                    'CANCELA LOS CAMBIOS Y DEJA LOS DATOS COMO ESTABAN
                    .CancelBatch
                    .MoveLast
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = False
                    botones
            'BORRAR
            ElseIf Index = 4 Then
            
                If (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
                    MsgBox "No Hay Ningun Registro Activo", vbOKOnly + vbInformation, "Informacion"
                Else
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        .Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        .MoveLast
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If .EOF Then
                        .MoveLast
                        If Err = 3021 Then
                            mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                        End If
                    End If
                End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        End With
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    With Adodc1
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                .RecordSource = ("Select * from BodegasProductoTerminado where CodigoBodega Like '" & TxtBuscar.Text & "%'")
                .Refresh
                DataGrid1.Refresh
            ElseIf OptNombre.Value = True Then
                .RecordSource = ("Select * from BodegasProductoTerminado where Descripcion like '" & TxtBuscar.Text & "*'")
                .Refresh
                DataGrid1.Refresh
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From BodegasProductoTerminado"
                .Refresh
                DataGrid1.Refresh
        End If
    End With
        TabBodegas.Tab = 1
End Sub

Private Sub datagrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.RecordSource = ("Select * from BodegasProductoTerminado order by " & DataGrid1.Columns(ColIndex).DataField)
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Form_Load()
'    Adodc1.DatabaseName = BasedeDatos
    'With Adodc1
    '    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.00;" & "Data Source=" & BasedeDatos & ";"
    '    .RecordSource = "BodegasProductoTerminado"
    'End With
End Sub

Private Sub OptCodigo_Click()
    Lbletiqueta.Caption = "Codigo"
End Sub

Private Sub OptNombre_Click()
    Lbletiqueta.Caption = "Descripcion"
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
End Sub
