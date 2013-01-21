VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DesperdicioPacas 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura De Desperdicio Pacas"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "DesperdicioPacas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabBodegas 
      Height          =   4695
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "DesperdicioPacas.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePacas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "DesperdicioPacas.frx":075C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridPacas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "DesperdicioPacas.frx":0BAE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DtpFecFin"
      Tab(2).Control(1)=   "DtpFecIni"
      Tab(2).Control(2)=   "CmdBuscar(1)"
      Tab(2).Control(3)=   "CmdBuscar(0)"
      Tab(2).Control(4)=   "Label1(2)"
      Tab(2).Control(5)=   "Label1(1)"
      Tab(2).ControlCount=   6
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -71880
         TabIndex        =   29
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64290819
         CurrentDate     =   37501
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -73560
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64290819
         CurrentDate     =   37501
      End
      Begin MSDBGrid.DBGrid DBGridPacas 
         Bindings        =   "DesperdicioPacas.frx":1000
         Height          =   3855
         Left            =   -74880
         OleObjectBlob   =   "DesperdicioPacas.frx":1018
         TabIndex        =   27
         Top             =   720
         Width           =   8175
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "DesperdicioPacas.frx":2613
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "DesperdicioPacas.frx":291D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Frame FramePacas 
         Caption         =   "Datos de Pacas"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   8115
         Begin MSMask.MaskEdBox MskTotal 
            Height          =   735
            Left            =   4080
            TabIndex        =   33
            Top             =   2760
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1296
            _Version        =   393216
            Appearance      =   0
            BackColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Fecha"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   0
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataPacas"
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "NoPaca"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   1
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "DesperdicioProceso"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "DesperdicioProveedor"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "Refil"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "ProductoNoConformeInterno"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   5
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "ProductoNoConformeExterno"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   6
            Top             =   2520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            DataField       =   "UnidadMedida"
            DataSource      =   "DataPacas"
            Height          =   285
            Index           =   7
            Left            =   2400
            TabIndex        =   7
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4080
            TabIndex        =   34
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidad De Medida"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto No Conforme Interno"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   2160
         End
         Begin VB.Label Label2 
            Caption         =   "Refil Y Otros"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   3240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto No Conforme Externo"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   2520
            Width           =   2205
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desperdicio Proveedor"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1620
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desperdicio Proceso"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "No. Paca"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   2
         Left            =   -71880
         TabIndex        =   31
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Index           =   1
         Left            =   -73560
         TabIndex        =   30
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.Data DataPacas 
      BackColor       =   &H80000014&
      Caption         =   "Captura Desperdicio Pacas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DesperdicioPacas"
      Top             =   5760
      Width           =   8115
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6840
      MouseIcon       =   "DesperdicioPacas.frx":2D5F
      Picture         =   "DesperdicioPacas.frx":31A1
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5520
      MouseIcon       =   "DesperdicioPacas.frx":5213
      Picture         =   "DesperdicioPacas.frx":5655
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "DesperdicioPacas.frx":5B87
      Picture         =   "DesperdicioPacas.frx":5FC9
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2880
      MouseIcon       =   "DesperdicioPacas.frx":64FB
      Picture         =   "DesperdicioPacas.frx":693D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "DesperdicioPacas.frx":6E6F
      Picture         =   "DesperdicioPacas.frx":72B1
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "DesperdicioPacas.frx":77E3
      Picture         =   "DesperdicioPacas.frx":7C25
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1200
   End
End
Attribute VB_Name = "DesperdicioPacas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VUltimaUnidadMedida As String
Dim RBuscaPaca As Recordset

Sub botones()
    If Bandera = True Then
         FramePacas.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         Msk.Item(0).SetFocus
         DataPacas.Visible = False
         DBGridPacas.Visible = False
    Else
         FramePacas.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataPacas.Visible = True
         DBGridPacas.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataPacas.Recordset
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
                    Msk.Item(0).Text = Date
                    Msk.Item(0).SetFocus
                    Msk.Item(7).Text = VUltimaUnidadMedida
                    TxtUsuario.Text = GUsuario
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
                    Msk.Item(0).SetFocus
                    TxtUsuario.Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                    VUltimaUnidadMedida = Msk.Item(7).Text
                    
                    'Fecha
                    If Not IsDate(Msk.Item(0).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(0).SetFocus
                        Exit Sub
                    End If
                    
                    
                    'No. Paca
                    If Not IsNumeric(Msk.Item(1).Text) Then
                        MsgBox "No. De Paca Debe Ser Numerico", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(1).SetFocus
                        Exit Sub
                    End If
                    
                    'Desperdicio Proceso
                    If Not IsNumeric(Msk.Item(2).Text) Then
                        MsgBox "Desperdicio De Proceso", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(2).SetFocus
                        Exit Sub
                    End If
                    
                    'Desperdicio Proveedor
                    If Not IsNumeric(Msk.Item(3).Text) Then
                        MsgBox "Desperdicio De Proveedor", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(3).SetFocus
                        Exit Sub
                    End If
                    
                    'Perfil
                    If Not IsNumeric(Msk.Item(4).Text) Then
                        MsgBox "Perfil Debe Ser Numerico", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(4).SetFocus
                        Exit Sub
                    End If
                    
                    'Producto No Conforme Interno
                    If Not IsNumeric(Msk.Item(5).Text) Then
                        MsgBox "Producto No Conforme Interno Debe Ser Numerico", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(5).SetFocus
                        Exit Sub
                    End If
                    
                    'Producto No Conforme Externo
                    If Not IsNumeric(Msk.Item(6).Text) Then
                        MsgBox "Producto No Conforme Externo Debe Ser Numerico", vbOKOnly + vbExclamation, "Verifique"
                        Msk.Item(6).SetFocus
                        Exit Sub
                    End If
                    
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "No. De Paca Con Esta Fecha ya existe", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(0).SetFocus
                        Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                     ElseIf Err <> 3022 And Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Msk.Item(0).SetFocus
                        Exit Sub
                     End If
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
            'CANCELAR
            ElseIf Index = 3 Then
                    'CANCELA LOS CAMBIOS Y DEJA LOS DATOS COMO ESTABAN
                    .CancelUpdate
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Msk.Item(1).SetFocus
                        Exit Sub
                    End If
                    Bandera = False
                    botones
            'BORRAR
            ElseIf Index = 4 Then
                                                     
            
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        DataPacas.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataPacas.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataPacas.Recordset.EOF Then
                        DataPacas.Recordset.MoveLast
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
    With DataPacas
                .RecordSource = ("Select * from DesperdicioPacas Where Fecha >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#")
                .Refresh
                DBGridPacas.Refresh
    End With
        TabBodegas.Tab = 1
End Sub

Private Sub DataPacas_Reposition()
        MskTotal.Text = Val(Msk.Item(2).Text) + Val(Msk.Item(3).Text) + Val(Msk.Item(4).Text) + Val(Msk.Item(5).Text) + Val(Msk.Item(6).Text)
End Sub

Private Sub DataPacas_Validate(Action As Integer, Save As Integer)
        MskTotal.Text = Val(Msk.Item(2).Text) + Val(Msk.Item(3).Text) + Val(Msk.Item(4).Text) + Val(Msk.Item(5).Text) + Val(Msk.Item(6).Text)
End Sub

Private Sub dbgridpacas_HeadClick(ByVal ColIndex As Integer)
    DataPacas.RecordSource = ("Select * from DesperdicioPacas order by " & DBGridPacas.Columns(ColIndex).DataField)
    DataPacas.Refresh
    DBGridPacas.Refresh
End Sub

Private Sub Form_Load()
    DataPacas.ConnectionString = GTipoProveedor
    
    DataPacas.Refresh
End Sub

Private Sub Msk_Change(Index As Integer)
        MskTotal.Text = Val(Msk.Item(2).Text) + Val(Msk.Item(3).Text) + Val(Msk.Item(4).Text) + Val(Msk.Item(5).Text) + Val(Msk.Item(6).Text)
End Sub

Private Sub Msk_GotFocus(Index As Integer)
        Msk.Item(Index).SelStart = 0
        Msk.Item(Index).SelLength = Len(Msk.Item(Index).Text)
End Sub

Private Sub Msk_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TabBodegas_Click(PreviousTab As Integer)
        If TabBodegas.Tab = 2 Then
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
        End If
End Sub

