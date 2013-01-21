VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FichaTecnicaConMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion De Materia Prima A Ficha Tecnica De Producto Terminado"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "FichaTecnicaConMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
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
      Height          =   6855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5655
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "FichaTecnicaConMateriaPrima.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":237C
      Picture         =   "FichaTecnicaConMateriaPrima.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Ultimo Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":2CF0
      Picture         =   "FichaTecnicaConMateriaPrima.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Siguiente Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":3664
      Picture         =   "FichaTecnicaConMateriaPrima.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Registro Anterior"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":3FD8
      Picture         =   "FichaTecnicaConMateriaPrima.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Primer Registro"
      Top             =   5760
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   5535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "FichaTecnicaConMateriaPrima.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameFichaTecnicaConMateriaPrima"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "FichaTecnicaConMateriaPrima.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "FichaTecnicaConMateriaPrima.frx":50B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBuscar(1)"
      Tab(2).Control(1)=   "CmdBuscar(0)"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "Lbletiqueta"
      Tab(2).ControlCount=   5
      Begin MSDataGridLib.DataGrid DbGrid1 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   31
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Top             =   2400
         Width           =   8175
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "FichaTecnicaConMateriaPrima.frx":550A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "FichaTecnicaConMateriaPrima.frx":5814
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   11
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2760
         Width           =   2085
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
         TabIndex        =   24
         Top             =   960
         Width           =   3165
         Begin VB.OptionButton OptFichaTecnica 
            Caption         =   "Ficha Tecnica"
            Height          =   225
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptMateriaPrima 
            Caption         =   "Materia Prima"
            Height          =   195
            Left            =   1680
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
         End
      End
      Begin VB.Frame FrameFichaTecnicaConMateriaPrima 
         Caption         =   "Datos Tipo Entrada Materia Prima"
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
         TabIndex        =   21
         Top             =   720
         Width           =   8115
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidad De Medida"
            Height          =   195
            Index           =   2
            Left            =   2880
            TabIndex        =   30
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Consumo Por Unidad"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   975
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
            Left            =   2880
            TabIndex        =   27
            Top             =   720
            Width           =   5175
         End
         Begin VB.Label LblFichaTecnica 
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
            Left            =   2880
            TabIndex        =   26
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label2 
            Caption         =   "Materia Prima"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Left            =   -70800
         TabIndex        =   25
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6240
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":5C56
      Picture         =   "FichaTecnicaConMateriaPrima.frx":6098
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   4920
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":810A
      Picture         =   "FichaTecnicaConMateriaPrima.frx":854C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   3600
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":8A7E
      Picture         =   "FichaTecnicaConMateriaPrima.frx":8EC0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2280
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":93F2
      Picture         =   "FichaTecnicaConMateriaPrima.frx":9834
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "FichaTecnicaConMateriaPrima.frx":9D66
      Picture         =   "FichaTecnicaConMateriaPrima.frx":A1A8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1200
   End
End
Attribute VB_Name = "FichaTecnicaConMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BFichaTecnica As Boolean
Dim BMateriaPrima As Boolean

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBuscaMateriasPrimas As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RFicha As New ADODB.Recordset
Dim VTexto As String

Sub botones()
    If Bandera = True Then
         FrameFichaTecnicaConMateriaPrima.Enabled = True
         CmdBotones.Item(0).Enabled = False
         
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         FrameOpciones.Visible = False
         DbGrid1.Visible = False
    Else
         FrameFichaTecnicaConMateriaPrima.Enabled = False
         CmdBotones.Item(0).Enabled = True
         
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
 
         FrameOpciones.Visible = True
         DbGrid1.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        
            If Index = 0 Then
                    Bandera = True
                    botones
                    Limpia_Campos
                    TxtTexto.Item(0).SetFocus
            'GRABAR
            ElseIf Index = 2 Then
                            VTexto = "'" & TxtTexto.Item(0).Text & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', " 'MATERIA PRIMA
                            VTexto = VTexto & TxtTexto.Item(2).Text & ", '" 'CONSUMO
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', " 'UNIDAD DE MEDIDA
                            VTexto = VTexto & "0" 'NOMBRE COMERCIAL
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into FichaTecnicaConMateriaPrima Values(" & VTexto & ")"
                            
                                'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Ficha Tecnica y Materia Prima Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Ficha Tecnica y Materia Prima Ya Existe", vbOKOnly + vbInformation, "Informacion"
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
                        RFicha.Requery
                        RFicha.MoveLast
                        Llena_Campos
                        CmdBotones.Item(0).SetFocus
            'CANCELAR
            ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RFicha.Delete
                        
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
                        RFicha.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RFicha.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            'Err.Clear
                        End If
                        
                        Llena_Campos
                        
                    End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
    
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RFicha.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RFicha.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RFicha.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RFicha.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RFicha.BOF Then
        RFicha.MoveFirst
    ElseIf RFicha.EOF Then
        RFicha.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
    TxtDatos.Text = ""
    Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
    If RBuscaMateriasPrimas.RecordCount > 0 Then
            Do Until RBuscaMateriasPrimas.EOF
                TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                RBuscaMateriasPrimas.MoveNext
            Loop
    End If
    
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RFicha = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptFichaTecnica.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RFicha, "Select * from FichaTecnicaConMateriaPrima where Esp_Tec like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RFicha, "Select * from FichaTecnicaConMateriaPrima where UPPER(Esp_Tec) like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptMateriaPrima.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RFicha, "Select * from FichaTecnicaConMateriaPrima where CodigoMateriaPrima like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RFicha, "Select * from FichaTecnicaConMateriaPrima where UPPER(CodigoMateriaPrima) like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RFicha, "Select * From FichaTecnicaConMateriaPrima")
        End If
        
        Set DbGrid1.DataSource = RFicha
    
        TabBodegas.Tab = 1
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DbGridBusqueda_DblClick()
            If BFichaTecnica = True Then
                TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(0).SetFocus
            Else
                TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(1).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DbGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BFichaTecnica = True Then
                    TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0).Text
                    TxtTexto.Item(0).SetFocus
                Else
                    TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
                    TxtTexto.Item(1).SetFocus
                End If
            End If
                FrameBusqueda.Visible = False

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
        RFicha.Sort = RFicha.Fields(ColIndex).Name
End Sub

Private Sub Form_Activate()
    TxtDatos.Text = ""
    Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
    If RBuscaMateriasPrimas.RecordCount > 0 Then
            Do Until RBuscaMateriasPrimas.EOF
                TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                RBuscaMateriasPrimas.MoveNext
            Loop
    End If
    
End Sub

Private Sub Form_Load()
        Set RFicha = New ADODB.Recordset
        Call Abrir_Recordset(RFicha, "Select * From FichaTecnicaConMateriaPrima")
        Set DbGrid1.DataSource = RFicha
        Llena_Campos
        
        'PARA HABILITAR EL GRID SOLO A USUARIOS AVANZADOS
        If GEditar = True Then
            DbGrid1.AllowAddNew = True
            DbGrid1.AllowUpdate = True
        End If
End Sub

Private Sub OptFichaTecnica_Click()
        Lbletiqueta.Caption = "Ficha Tecnica"
End Sub



Private Sub OptMateriaPrima_Click()
        Lbletiqueta.Caption = "Materia Prima"
End Sub

Private Sub TabBodegas_Click(PreviousTab As Integer)
        If TabBodegas.Tab = 0 Then
            CmdBotones.Item(4).Enabled = True
            If CmdBotones.Item(2).Enabled = False Then
                Llena_Campos
            End If
        Else
            CmdBotones.Item(4).Enabled = False
        End If
        
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


Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
    'MATERIA PRIMA
    If BMateriaPrima = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    'FICHA TECNICA
    ElseIf BFichaTecnica = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    End If
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"

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
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
            Set RBuscaFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFichaTecnica.Caption = ""
                End If
        'BUSCA TODAS LAS MATERIAS PRIMAS QUE TIENE ASIGNADA LA FICHA TECNICA
        TxtDatos.Text = ""
        Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
        If RBuscaMateriasPrimas.RecordCount > 0 Then
                Do Until RBuscaMateriasPrimas.EOF
                    TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                    RBuscaMateriasPrimas.MoveNext
                Loop
        End If
        
        'BUSCA LA DESCRIPCION DE MATERIA PRIMA
        ElseIf Index = 1 Then
           Set RBuscaMateriaPrima = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip, UnidadMedida From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip, UnidadMedida From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                End If
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    If IsNull(RBuscaMateriaPrima!Descrip) Then
                        LblMateriaPrima.Caption = ""
                    Else
                        LblMateriaPrima.Caption = RBuscaMateriaPrima!Descrip
                    End If
                    If IsNull(RBuscaMateriaPrima!unidadMedida) Then
                        TxtTexto.Item(3).Text = ""
                    Else
                        TxtTexto.Item(3).Text = RBuscaMateriaPrima!unidadMedida
                    End If
                Else
                    LblMateriaPrima.Caption = ""
                    TxtTexto.Item(3).Text = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 0 Or Index = 1 Then
            Set RBusqueda = New ADODB.Recordset
            'SI ELIGE FICHA TECNICA
            If Index = 0 Then
                BFichaTecnica = True
                BMateriaPrima = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            'SI ELIGE MATERIA PRIMA
            Else
                BFichaTecnica = False
                BMateriaPrima = True
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            End If
        
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "3000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
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
    
    If KeyAscii = 43 Then
      If Index = 0 Or Index = 1 Then
            Set RBusqueda = New ADODB.Recordset
            'SI ELIGE FICHA TECNICA
            If Index = 0 Then
                BFichaTecnica = True
                BMateriaPrima = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            'SI ELIGE MATERIA PRIMA
            Else
                BFichaTecnica = False
                BMateriaPrima = True
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            End If
        
            
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "3000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
        End If
    End If
    
End Sub
Public Sub Llena_Campos()
On Error Resume Next
        TxtTexto.Item(0).Text = RFicha!Esp_Tec
        TxtTexto.Item(1).Text = RFicha!CodigoMateriaPrima
        If IsNull(RFicha!Consumo) Then
            TxtTexto.Item(2).Text = "0"
        Else
            TxtTexto.Item(2).Text = RFicha!Consumo
        End If
        If IsNull(RFicha!UnidadDeMedida) Then
            TxtTexto.Item(3).Text = ""
        Else
            TxtTexto.Item(3).Text = RFicha!UnidadDeMedida
        End If
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
End Sub

