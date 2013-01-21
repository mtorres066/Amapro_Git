VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Clientes 
   BackColor       =   &H00008000&
   Caption         =   "Clientes"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   Icon            =   "Clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1080
      MouseIcon       =   "Clientes.frx":08CA
      Picture         =   "Clientes.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2280
      MouseIcon       =   "Clientes.frx":123E
      Picture         =   "Clientes.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3480
      MouseIcon       =   "Clientes.frx":1BB2
      Picture         =   "Clientes.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4680
      MouseIcon       =   "Clientes.frx":2526
      Picture         =   "Clientes.frx":2968
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5880
      MouseIcon       =   "Clientes.frx":2E9A
      Picture         =   "Clientes.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   7080
      MouseIcon       =   "Clientes.frx":380E
      Picture         =   "Clientes.frx":3C50
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   8760
      MouseIcon       =   "Clientes.frx":5CC2
      Picture         =   "Clientes.frx":6104
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Ultimo Registro"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   8400
      MouseIcon       =   "Clientes.frx":6636
      Picture         =   "Clientes.frx":6A78
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Siguiente Registro"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "Clientes.frx":6FAA
      Picture         =   "Clientes.frx":73EC
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Registro Anterior"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "Clientes.frx":791E
      Picture         =   "Clientes.frx":7D60
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Primer Registro"
      Top             =   5160
      Width           =   375
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
      BackColor       =   32768
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Clientes.frx":8292
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameClientes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Clientes.frx":85AC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridClientes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "Clientes.frx":89FE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGridClientes 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   31
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "CodigoCliente"
            Caption         =   "Codigo"
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
            DataField       =   "Fax"
            Caption         =   "Fax"
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
            DataField       =   "Nit"
            Caption         =   "Nit"
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
         BeginProperty Column06 
            DataField       =   "Contacto"
            Caption         =   "Contacto"
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
         BeginProperty Column07 
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
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   8775
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccion o Busqueda"
            Height          =   855
            Index           =   6
            Left            =   6480
            Picture         =   "Clientes.frx":8E50
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Todos"
            Height          =   855
            Index           =   7
            Left            =   6480
            Picture         =   "Clientes.frx":9292
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3000
            Width           =   2055
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   17
            Top             =   1680
            Width           =   2535
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
            Picture         =   "Clientes.frx":959C
            Style           =   1  'Graphical
            TabIndex        =   16
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
            Picture         =   "Clientes.frx":99DE
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Left            =   3840
            TabIndex        =   28
            Top             =   1680
            Width           =   2055
         End
      End
      Begin VB.Frame FrameClientes 
         Caption         =   "Datos del Cliente"
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
         Height          =   3735
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   8655
         Begin VB.ComboBox CboEst 
            Height          =   315
            ItemData        =   "Clientes.frx":9CE8
            Left            =   1320
            List            =   "Clientes.frx":9CF2
            TabIndex        =   36
            Text            =   "ACTIVO"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   6
            Top             =   2040
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   7
            Top             =   2400
            Width           =   6240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1680
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1320
            Width           =   3240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   3
            Top             =   960
            Width           =   6240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   2
            Top             =   600
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   1
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   37
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   30
            Top             =   3120
            Width           =   540
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nit"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   29
            Top             =   2040
            Width           =   195
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Encargado"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   26
            Top             =   2400
            Width           =   780
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   25
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   24
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   23
            Top             =   960
            Width           =   675
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   22
            Top             =   600
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
            TabIndex        =   21
            Top             =   240
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer
Dim BEditar As Boolean
Dim VTexto As String
Dim RClientes As New ADODB.Recordset

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    
    
        'AGREGAR
        If Index = 0 Then
                TabDepartamentos.Tab = 0
                Bandera = True
                botones
                Limpia_Campos
                'HABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = True
                TxtUsuario.Text = GUsuario
                TxtTexto.Item(0).SetFocus
                BEditar = False
        'EDITAR
        ElseIf Index = 1 Then
                TabDepartamentos.Tab = 0
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = False
                TxtTexto.Item(1).SetFocus
                TxtUsuario.Text = GUsuario
                BEditar = True
        'GRABAR
        ElseIf Index = 2 Then
                        
                    If (CboEst.Text <> "ACTIVO" And CboEst.Text <> "INACTIVO") Then
                        MsgBox "Estado solo puede ser ACTIVO o INACTIVO"
                        CboEst.SetFocus
                        Exit Sub
                    End If
                    
                        
                    'AGREGAR
                    If BEditar = False Then
                            VTexto = "'" & TxtTexto.Item(0).Text & "', '" 'CODIGO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'NOMBRE
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'DIRCCION
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '" 'TELEFONO
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'FAX
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', '" 'NIT
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', '" 'CONTACTO
                            VTexto = VTexto & TxtUsuario.Text & "', '" 'USUARIO
                            VTexto = VTexto & CboEst.Text & "'" '
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Clientes Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            VTexto = "Descripcion = '" & TxtTexto.Item(1).Text & "', " 'NOMBRE
                            VTexto = VTexto & "Direccion = '" & TxtTexto.Item(2).Text & "', " 'DIRCCION
                            VTexto = VTexto & "Telefono = '" & TxtTexto.Item(3).Text & "', " 'TELEFONO
                            VTexto = VTexto & "Fax = '" & TxtTexto.Item(4).Text & "', " 'FAX
                            VTexto = VTexto & "Nit = '" & TxtTexto.Item(5).Text & "', " 'NIT
                            VTexto = VTexto & "Contacto = '" & TxtTexto.Item(6).Text & "', " 'CONTACTO
                            VTexto = VTexto & "Usuario = '" & TxtUsuario.Text & "', " 'USUARIO
                            VTexto = VTexto & "Estado = '" & CboEst.Text & "'" '
                            VTexto = VTexto & " Where CodigoCliente = '" & TxtTexto.Item(0).Text & "'" 'CODIGO
                            
                            Conexion.Execute "UPDATE Clientes SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo Cliente Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo Cliente Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                        
                        'HABILITA LA LLAVE
                        TxtTexto.Item(0).Enabled = True
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RClientes.Requery
                        RClientes.MoveLast
                        Llena_Campos

        'CANCELAR
        ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    'HABILITA LA LLAVE
                    TxtTexto.Item(0).Enabled = True
                    
        ElseIf Index = 4 Then ' BORRAR
        
            On Error Resume Next
                VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If VMensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RClientes.Delete
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RClientes.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RClientes.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    Set RClientes = New ADODB.Recordset
                    If OptBusqueda.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RClientes, "Select * From Clientes Where CodigoCliente Like '" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RClientes, "Select * From Clientes Where UPPER(CodigoCliente) Like '" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RClientes, "Select * From Clientes Where Descripcion Like '" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RClientes, "Select * From Clientes Where UPPER(Descripcion) Like '" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                    End If
                    Set DGridClientes.DataSource = RClientes
                    TabDepartamentos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    Set RClientes = New ADODB.Recordset
                    Call Abrir_Recordset(RClientes, "Select * From Clientes")
                    Set DGridClientes.DataSource = RClientes
                    TabDepartamentos.Tab = 1
        End If

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameClientes.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         DGridClientes.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameClientes.Enabled = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         DGridClientes.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub



Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RClientes.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RClientes.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RClientes.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RClientes.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RClientes.BOF Then
        RClientes.MoveFirst
    ElseIf RClientes.EOF Then
        RClientes.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub


Private Sub DGridClientes_HeadClick(ByVal ColIndex As Integer)
            RClientes.Sort = RClientes.Fields(ColIndex).Name
    

End Sub

Private Sub Form_Load()
        Set RClientes = New ADODB.Recordset
        Call Abrir_Recordset(RClientes, "Select * From Clientes")
        Set DGridClientes.DataSource = RClientes
        Llena_Campos

End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Codigo"
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Descripcion"
    End If
    
End Sub

Private Sub TabDepartamentos_Click(PreviousTab As Integer)
        If TabDepartamentos.Tab = 0 Then
            CmdBotones.Item(4).Enabled = True
            If CmdBotones.Item(2).Enabled = False Then
                Llena_Campos
            End If
        Else
            CmdBotones.Item(4).Enabled = False
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

Private Sub TxtTexto_GotFocus(Index As Integer)
        TxtTexto.Item(Index).SelStart = 0
        TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
                SendKeys "{tab}"
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'CODIGO
            If IsNull(RClientes!CodigoCliente) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RClientes!CodigoCliente
            End If
        'DESCRIPCION
            If IsNull(RClientes!Descripcion) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RClientes!Descripcion
            End If
        'DIRECCION
            If IsNull(RClientes!Direccion) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RClientes!Direccion
            End If
        'TELEFONO
            If IsNull(RClientes!Telefono) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RClientes!Telefono
            End If
        'FAX
            If IsNull(RClientes!Fax) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RClientes!Fax
            End If
        'NIT
            If IsNull(RClientes!Nit) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RClientes!Nit
            End If
        'CONTACTO
            If IsNull(RClientes!Contacto) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = RClientes!Contacto
            End If
        'USUARIO
            If IsNull(RClientes!Usuario) Then
                TxtUsuario.Text = ""
            Else
                TxtUsuario.Text = RClientes!Usuario
            End If
            
            If IsNull(RClientes!Estado) Then
                CboEst.Text = ""
            Else
                CboEst.Text = RClientes!Estado
            End If
            
        
        If Err <> 0 Then
            MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
        

        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(6).Text = ""
        CboEst.Text = "ACTIVO"
        TxtUsuario.Text = ""
        
End Sub

