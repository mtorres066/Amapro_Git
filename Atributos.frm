VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Atributos 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento De Atributos"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   Icon            =   "Atributos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "Atributos.frx":08CA
      Picture         =   "Atributos.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Primer Registro"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "Atributos.frx":123E
      Picture         =   "Atributos.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Registro Anterior"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   9840
      MouseIcon       =   "Atributos.frx":1BB2
      Picture         =   "Atributos.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Siguiente Registro"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   10200
      MouseIcon       =   "Atributos.frx":2526
      Picture         =   "Atributos.frx":2968
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ultimo Registro"
      Top             =   4080
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Atributos.frx":2E9A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAtributos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Atributos.frx":31B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Atributos.frx":3606
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBuscar(0)"
      Tab(2).Control(1)=   "CmdBuscar(1)"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "Lbletiqueta"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -67800
         Picture         =   "Atributos.frx":3A58
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -67800
         Picture         =   "Atributos.frx":3E9A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2640
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5318
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Codigo"
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
            DataField       =   "Menores"
            Caption         =   "Menores"
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
            DataField       =   "Mayores"
            Caption         =   "Mayores"
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
            DataField       =   "Criticos"
            Caption         =   "Criticos"
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
            DataField       =   "Muestra"
            Caption         =   "Muestra"
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
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -71760
         TabIndex        =   18
         ToolTipText     =   " "
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Frame FrameAtributos 
         Caption         =   "Datos del Atributo"
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
         Height          =   2295
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   10275
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Codigo"
            DataSource      =   "DataAtributos"
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtMen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Menores"
            DataSource      =   "DataAtributos"
            Height          =   285
            Left            =   1080
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox TxtMay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Mayores"
            DataSource      =   "DataAtributos"
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox TxtCri 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Criticos"
            DataSource      =   "DataAtributos"
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox TxtMue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Muestra"
            DataSource      =   "DataAtributos"
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Menores"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Mayores"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Criticos"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Muestra"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   735
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
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   8400
      MouseIcon       =   "Atributos.frx":41A4
      Picture         =   "Atributos.frx":45E6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   6960
      MouseIcon       =   "Atributos.frx":6658
      Picture         =   "Atributos.frx":6A9A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   5400
      MouseIcon       =   "Atributos.frx":6FCC
      Picture         =   "Atributos.frx":740E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1515
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3840
      MouseIcon       =   "Atributos.frx":7940
      Picture         =   "Atributos.frx":7D82
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1515
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   2400
      MouseIcon       =   "Atributos.frx":82B4
      Picture         =   "Atributos.frx":86F6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   960
      MouseIcon       =   "Atributos.frx":8C28
      Picture         =   "Atributos.frx":906A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1400
   End
End
Attribute VB_Name = "Atributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEditar As Boolean

Dim RAtributos As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameAtributos.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         
         DataGrid1.Visible = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
         CmdBuscar.Item(0).Visible = False
         CmdBuscar.Item(1).Visible = False
    Else
         FrameAtributos.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         
         DataGrid1.Visible = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         CmdBuscar.Item(0).Visible = True
         CmdBuscar.Item(1).Visible = True
    End If
End Sub

Private Sub CmdAgregar_Click()
        Bandera = True
        botones
        Limpia_Campos
        'HABILITA EL CODIGO DE NUEVO
        TxtCod.Enabled = True
        TxtCod.SetFocus
        BEditar = False
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                'BORRA EL REGISTRO
                RAtributos.Delete
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
                        RAtributos.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RAtributos.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
            End If
            
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RAtributos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RAtributos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RAtributos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RAtributos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RAtributos.BOF Then
        RAtributos.MoveFirst
    ElseIf RAtributos.EOF Then
        RAtributos.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RAtributos = New ADODB.Recordset
                    
        If Index = 0 Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RAtributos, "Select * from Atributos where Codigo like '" & TxtBuscar.Text & "%'")
                    Else
                        Call Abrir_Recordset(RAtributos, "Select * from Atributos where UPPER(Codigo) like '" & UCase(TxtBuscar.Text) & "%'")
                    End If
        Else
                    Call Abrir_Recordset(RAtributos, "Select * from Atributos")
        End If
                    
                    Set DataGrid1.DataSource = RAtributos
                    
                    SSTab1.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        Llena_Campos
        'HABILITA EL CODIGO DE NUEVO
        TxtCod.Enabled = True
End Sub

Private Sub CmdEditar_Click()
        Bandera = True
        botones
        BEditar = True
        TxtCod.Enabled = False
        TxtMen.SetFocus
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
             'MENORES
             If Not IsNumeric(TxtMen.Text) Then
                  MsgBox "Menores Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                  Exit Sub
             End If
             'MAYORES
             If Not IsNumeric(TxtMay.Text) Then
                  MsgBox "Mayores Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                  Exit Sub
             End If
             'CRITICOS
             If Not IsNumeric(TxtCri.Text) Then
                  MsgBox "Criticos Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                  Exit Sub
             End If
             'MUESTRA
             If Not IsNumeric(TxtMue.Text) Then
                  MsgBox "Muestra Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                  Exit Sub
             End If
                 
             If BEditar = False Then 'ESTA AGREGANDO UN REGISTRO
                  Conexion.Execute "Insert Into Atributos Values('" & TxtCod.Text & "'," & TxtMen.Text & "," & TxtMay.Text & "," & TxtCri.Text & "," & TxtMue.Text & ")"
             Else 'ESTA EDITANDO UN REGISTRO
                  Conexion.Execute "UPDATE Atributos SET Menores = " & TxtMen.Text & ", Mayores = " & TxtMay.Text & ", Criticos = " & TxtCri.Text & ", Muestra = " & TxtMue.Text & " Where Codigo = '" & TxtCod.Text & "'"
             End If
             
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo De Grupo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtCod.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo De Grupo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtCod.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
             
                Bandera = False
                botones
                'HABILITA EL CODIGO DE NUEVO
                TxtCod.Enabled = True
                'VUELVE A DIBUJAR EL RECORDSET Y LLENA LOS CAMPOS
                RAtributos.Requery
            

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
        If RAtributos.RecordCount > 0 Then
            RAtributos.Sort = RAtributos.Fields(ColIndex).Name
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If
        End If
    
End Sub

Private Sub Form_Load()
        Set RAtributos = New ADODB.Recordset
        Call Abrir_Recordset(RAtributos, "Select * From Atributos")
        Set DataGrid1.DataSource = RAtributos
        Llena_Campos
End Sub

Private Sub Form_Unload(Cancel As Integer)
        On Error Resume Next
            RAtributos.Close
            Set RAtributos = Nothing
            If Err <> 0 Then
            End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
        If SSTab1.Tab = 0 Then
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
            CmdBorrar.Enabled = True
        Else
            CmdBorrar.Enabled = False
        End If
End Sub


Private Sub TxtCod_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub txtmay_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub
Private Sub txtcri_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub TxtMen_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Public Sub Llena_Campos()
On Error Resume Next
        TxtCod.Text = RAtributos!Codigo
        TxtMen.Text = RAtributos!Menores
        TxtMay.Text = RAtributos!Mayores
        TxtCri.Text = RAtributos!Criticos
        TxtMue.Text = RAtributos!Muestra
        If Err <> 0 Then
            Err.Clear
        End If
End Sub

Public Sub Limpia_Campos()
        TxtCod.Text = ""
        TxtMen.Text = ""
        TxtMay.Text = ""
        TxtCri.Text = ""
        TxtMue.Text = ""
End Sub

Private Sub TxtMue_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub
