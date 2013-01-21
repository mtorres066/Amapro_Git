VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SolicitudCierre 
   BackColor       =   &H0080C0FF&
   Caption         =   "Autorizacion De Solicitudes"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "SolicitudCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11880
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
      Height          =   8055
      Left            =   960
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4092
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "SolicitudCierre.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   6855
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   12091
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
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "SolicitudCierre.frx":3D6C
      Left            =   3480
      List            =   "SolicitudCierre.frx":3D79
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtUbi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   19
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Empresa"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Departamento"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Seccion"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Maquina"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sistema"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptUbi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Equipo"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20971523
      CurrentDate     =   38483
   End
   Begin MSComCtl2.DTPicker DTPFecIni 
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20971523
      CurrentDate     =   38483
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Estado De Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton OptOpc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Todas"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton OptOpc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobada"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptOpc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Denegada"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptOpc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Pendiente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DbGrid 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   735
      Left            =   10800
      Picture         =   "SolicitudCierre.frx":3D9C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdVerDatos 
      Caption         =   "&Ver Datos"
      Height          =   735
      Left            =   9720
      Picture         =   "SolicitudCierre.frx":5E0E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LblUbi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Equipo"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label LblUbiDes 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Hasta"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Desde"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "SolicitudCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RVerDatos As New ADODB.Recordset
Dim RDetalleOrden As New ADODB.Recordset

Dim RBusqueda As New ADODB.Recordset

Dim RBuscaEmpresa As New ADODB.Recordset
Dim RBuscaDepartamento As New ADODB.Recordset
Dim RBuscaSeccion As New ADODB.Recordset
Dim RBuscaMaquina As New ADODB.Recordset
Dim RBuscaSistema As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset

Dim BEmpresa As Boolean
Dim BDepartamento As Boolean
Dim BSeccion As Boolean
Dim BMaquina As Boolean
Dim BSistema As Boolean
Dim BEquipo As Boolean


Dim VTexto As String


Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub CmdVerDatos_Click()
On Error Resume Next
MousePointer = 11
        Set RVerDatos = New ADODB.Recordset
                
                        If GOrigenDeDatos = "AmaproAccess" Then
                            VTexto = "Select ES.Documento, ES.Fecha, ES.Estado, ES.Observaciones, ES.Observaciones2, E.Descripcion, EM.Descripcion, D.Descripcion, S.Descripcion, M.Descripcion, SI.Descripcion, EQ.Descripcion, ES.Usuario From M_EncabezadoSolicitud ES, Empleados E, M_Empresas EM, M_Departamentos D, M_Secciones S, M_Maquinas M, M_Sistemas SI, M_Equipos EQ Where ES.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And ES.Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And ES.Empleado = E.Codigo And ES.Empresa = EM.Codigo And ES.Departamento = D.Codigo And ES.Seccion = S.Codigo And ES.Maquina = M.Codigo And ES.Sistema = Si.Codigo And ES.Equipo = EQ.Codigo"
                        Else 'ORACLE
                            VTexto = "Select ES.Documento, ES.Fecha, ES.Estado, ES.Observaciones, ES.Observaciones2, E.Descripcion, EM.Descripcion, D.Descripcion, S.Descripcion, M.Descripcion, SI.Descripcion, EQ.Descripcion, ES.Usuario From M_EncabezadoSolicitud ES, Empleados E, M_Empresas EM, M_Departamentos D, M_Secciones S, M_Maquinas M, M_Sistemas SI, M_Equipos EQ Where ES.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And ES.Fecha <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And ES.Empleado = E.Codigo And UPPER(ES.Empleado) = UPPER(E.Codigo) And UPPER(ES.Empresa) = UPPER(EM.Codigo) And UPPER(ES.Departamento) = UPPER(D.Codigo) And UPPER(ES.Seccion) = UPPER(S.Codigo) And UPPER(ES.Maquina) = UPPER(M.Codigo) And UPPER(ES.Sistema) = UPPER(Si.Codigo) And UPPER(ES.Equipo) = UPPER(EQ.Codigo)"
                        End If
                        
                        'OPCIONES
                        If OptUbi.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Empresa Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Empresa) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        ElseIf OptUbi.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Departamento Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Departamento) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        ElseIf OptUbi.Item(2).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Seccion Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Seccion) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        ElseIf OptUbi.Item(3).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Maquina Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Maquina) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        ElseIf OptUbi.Item(4).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Sistema Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Sistema) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        ElseIf OptUbi.Item(5).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " And ES.Equipo Like '" & TxtUbi.Text & "%'"
                            Else 'ORACLE
                                VTexto = VTexto & " And UPPER(ES.Equipo) Like '" & UCase(TxtUbi.Text) & "%'"
                            End If
                        End If
                        
                        If OptOpc.Item(0).Value = True Then
                            VTexto = VTexto & " And ES.Estado = 'PENDIENTE' Order by Documento"
                        ElseIf OptOpc.Item(1).Value = True Then
                            VTexto = VTexto & " And ES.Estado = 'DENEGADA' Order by Documento"
                        ElseIf OptOpc.Item(2).Value = True Then
                            VTexto = VTexto & " And ES.Estado = 'APROBADA' Order by Documento"
                        ElseIf OptOpc.Item(3).Value = True Then
                            VTexto = VTexto & " Order by Documento"
                        End If
                
                    Call Abrir_Recordset(RVerDatos, VTexto)
                    Set DbGrid.DataSource = RVerDatos
                    
                    DbGrid.Columns(0).Width = "800"
                    DbGrid.Columns(1).Width = "1000"
                    DbGrid.Columns(2).Width = "1500"
                    DbGrid.Columns(3).Width = "3500"
                    DbGrid.Columns(4).Width = "3500"
                    DbGrid.Columns(5).Width = "2000"
                    DbGrid.Columns(6).Width = "1000"
                    DbGrid.Columns(7).Width = "2000"
                    DbGrid.Columns(8).Width = "2000"
                    DbGrid.Columns(9).Width = "2000"
                    DbGrid.Columns(10).Width = "2000"
                    DbGrid.Columns(11).Width = "2000"
                    DbGrid.Columns(12).Width = "2000"
                                       
                    
                    DbGrid.Columns(1).NumberFormat = "dd/mm/yyyy"
                    
                    DbGrid.Columns(0).Caption = "# Solicitud"
                    DbGrid.Columns(1).Caption = "Fecha"
                    DbGrid.Columns(2).Caption = "Estado"
                    DbGrid.Columns(3).Caption = "Peticion"
                    DbGrid.Columns(4).Caption = "Respuesta"
                    DbGrid.Columns(5).Caption = "Empleado"
                    DbGrid.Columns(6).Caption = "Empresa"
                    DbGrid.Columns(7).Caption = "Departamento"
                    DbGrid.Columns(8).Caption = "Seccion"
                    DbGrid.Columns(9).Caption = "Maquina"
                    DbGrid.Columns(10).Caption = "Sistema"
                    DbGrid.Columns(11).Caption = "Equipo"
                    DbGrid.Columns(12).Caption = "Usuario"
                    
                    
                    DbGrid.Columns(2).Button = True
                    
                    DbGrid.Columns(0).Locked = True
                    DbGrid.Columns(1).Locked = True
                    DbGrid.Columns(2).Locked = False
                    DbGrid.Columns(3).Locked = True
                    DbGrid.Columns(4).Locked = False
                    DbGrid.Columns(5).Locked = True
                    DbGrid.Columns(6).Locked = True
                    DbGrid.Columns(7).Locked = True
                    DbGrid.Columns(8).Locked = True
                    DbGrid.Columns(9).Locked = True
                    DbGrid.Columns(10).Locked = True
                    DbGrid.Columns(11).Locked = True
                    DbGrid.Columns(12).Locked = True
                    
                    
                    
            MousePointer = 0
            
                If Err <> 0 Then
                    MsgBox "ERROR " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Err.Clear
                End If

End Sub

Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
                RVerDatos.Sort = RVerDatos.Fields(ColIndex).Name
                
End Sub

Private Sub DbGrid_Scroll(Cancel As Integer)
     '//Oculta la lista si hace Scroll
         List1.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
            TxtUbi.Text = DBGridBusqueda.Columns(0).Text
            TxtUbi.SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DBGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                            TxtUbi.Text = DBGridBusqueda.Columns(0).Text
                            TxtUbi.SetFocus
                            FrameBusqueda.Visible = False
            End If

End Sub

Private Sub DbGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
        '// Use este evento para que cuando el usuario teclee un caracter sobrela celda
        '// se despliegue la lista. Es decir, se obliga al usuario a usar un ítem de la lista.
        '// En caso de dar al usuario libertad de escribir, elimine las siguientes líneas (If-End If),
        '// o precedales con un comentario
            'If (ColIndex = 12 Or ColIndex = 13 Or ColIndex = 14 Or ColIndex = 15) Then
            If (ColIndex = 2) Then
               '// Se obliga a seleccionar de la lista:
               Cancel = True
               DbGrid_ButtonClick (ColIndex)
            End If

End Sub

Private Sub DbGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
            'ESTADO
            If ColIndex = 2 Then
                If DbGrid.Columns(2).Text <> "PENDIENTE" And DbGrid.Columns(2) <> "APROBADA" And DbGrid.Columns(2) <> "DENEGADA" Then
                    Cancel = True
                    MsgBox "Estado Incorrecto", vbOKOnly + vbInformation, "Informacion"
                End If
            
            End If
End Sub

Private Sub DbGrid_ButtonClick(ByVal ColIndex As Integer)
        'SI PRECIONA EL BUTON DE LA COLUMNA DE CALIDAD
        If ColIndex = 2 Then
            Dim C As Column
            Set C = DbGrid.Columns(ColIndex)
            With List1
                '// Despliegue de la lista al lado de la celda.
                '// Elimine los comentarios de las dos siguientes líneas
                '// y coloque comentarios a las tres posteriores. A su gusto
                
                .Left = DbGrid.Left + C.Left + C.Width
                .Top = DbGrid.Top + DbGrid.RowTop(DbGrid.Row)

                '// Lista debajo de la celda, al estilo ComboBox (3 líneas)
                .Left = DbGrid.Left + C.Left
                .Top = DbGrid.Top + DbGrid.RowTop(DbGrid.Row) + DbGrid.RowHeight
                .Width = C.Width + 15

                .ListIndex = 0
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
        End If

End Sub

Private Sub Form_Load()
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date
End Sub


Private Sub OptUbi_Click(Index As Integer)
        If Index = 0 Then
            LblUbi.Caption = "Empresa"
        ElseIf Index = 1 Then
            LblUbi.Caption = "Departamento"
        ElseIf Index = 2 Then
            LblUbi.Caption = "Seccion"
        ElseIf Index = 3 Then
            LblUbi.Caption = "Maquina"
        ElseIf Index = 4 Then
            LblUbi.Caption = "Sistema"
        ElseIf Index = 5 Then
            LblUbi.Caption = "Equipo"
        End If
            TxtUbi.SetFocus

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            If BEmpresa = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            ElseIf BDepartamento = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            ElseIf BSeccion = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            ElseIf BMaquina = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            ElseIf BSistema = True Then
                    'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            
            ElseIf BEquipo = True Then
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                            
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos where Codigo Like '%" & TxtBusqueda.Text & "%'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                                End If
                        End If
            End If
                    
                    If RBusqueda.RecordCount > 0 Then
                    End If
                    
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtUbi_Change()
            'EMPRESA
            If OptUbi.Item(0).Value = True Then
                    Set RBuscaEmpresa = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaEmpresa, "Select Descripcion From M_Empresas Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaEmpresa, "Select Descripcion From M_Empresas Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaEmpresa.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaEmpresa!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            'DEPARTAMENTO
            ElseIf OptUbi.Item(1).Value = True Then
                    Set RBuscaDepartamento = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From M_Departamentos Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From M_Departamentos Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaDepartamento.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaDepartamento!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            'SECCION
            ElseIf OptUbi.Item(2).Value = True Then
                    Set RBuscaSeccion = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaSeccion, "Select Descripcion From M_Secciones Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaSeccion, "Select Descripcion From M_Secciones Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaSeccion.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaSeccion!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            'MAQUINA
            ElseIf OptUbi.Item(3).Value = True Then
                    Set RBuscaMaquina = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaMaquina, "Select Descripcion From M_Maquinas Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaMaquina, "Select Descripcion From M_Maquinas Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaMaquina.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaMaquina!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            'SISTEMA
            ElseIf OptUbi.Item(4).Value = True Then
                    Set RBuscaSistema = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaSistema, "Select Descripcion From M_Sistemas Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaSistema, "Select Descripcion From M_Sistemas Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaSistema.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaSistema!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            'EQUIPO
            ElseIf OptUbi.Item(5).Value = True Then
                    Set RBuscaEquipo = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From M_Equipos Where Codigo = '" & TxtUbi.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From M_Equipos Where UPPER(Codigo) = '" & UCase(TxtUbi.Text) & "'")
                        End If
                        If RBuscaEquipo.RecordCount > 0 Then
                            LblUbiDes.Caption = RBuscaEquipo!Descripcion
                        Else
                            LblUbiDes.Caption = ""
                        End If
            End If

End Sub

Private Sub TxtUbi_DblClick()
            Set RBusqueda = New ADODB.Recordset
        
        If OptUbi.Item(0).Value = True Then
            BEmpresa = True
            BDepartamento = False
            BSeccion = False
            BMaquina = False
            BSistema = False
            BEquipo = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas")
        ElseIf OptUbi.Item(1).Value = True Then
            BEmpresa = False
            BDepartamento = True
            BSeccion = False
            BMaquina = False
            BSistema = False
            BEquipo = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos")
        ElseIf OptUbi.Item(2).Value = True Then
            BEmpresa = False
            BDepartamento = False
            BSeccion = True
            BMaquina = False
            BSistema = False
            BEquipo = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones")
        ElseIf OptUbi.Item(3).Value = True Then
            BEmpresa = False
            BDepartamento = False
            BSeccion = False
            BMaquina = True
            BSistema = False
            BEquipo = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas")
        ElseIf OptUbi.Item(4).Value = True Then
            BEmpresa = False
            BDepartamento = False
            BSeccion = False
            BMaquina = False
            BSistema = True
            BEquipo = False
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas")
        ElseIf OptUbi.Item(5).Value = True Then
            BEmpresa = False
            BDepartamento = False
            BSeccion = False
            BMaquina = False
            BSistema = False
            BEquipo = True
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos")
        End If
                
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus

End Sub

Private Sub TxtUbi_GotFocus()
        TxtUbi.SelStart = 0
        TxtUbi.SelLength = Len(TxtUbi.Text)
End Sub

Private Sub TxtUbi_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                    
                    Set RBusqueda = New ADODB.Recordset
        
                        If OptUbi.Item(0).Value = True Then
                            BEmpresa = True
                            BDepartamento = False
                            BSeccion = False
                            BMaquina = False
                            BSistema = False
                            BEquipo = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Empresas")
                        ElseIf OptUbi.Item(1).Value = True Then
                            BEmpresa = False
                            BDepartamento = True
                            BSeccion = False
                            BMaquina = False
                            BSistema = False
                            BEquipo = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Departamentos")
                        ElseIf OptUbi.Item(2).Value = True Then
                            BEmpresa = False
                            BDepartamento = False
                            BSeccion = True
                            BMaquina = False
                            BSistema = False
                            BEquipo = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Secciones")
                        ElseIf OptUbi.Item(3).Value = True Then
                            BEmpresa = False
                            BDepartamento = False
                            BSeccion = False
                            BMaquina = True
                            BSistema = False
                            BEquipo = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Maquinas")
                        ElseIf OptUbi.Item(4).Value = True Then
                            BEmpresa = False
                            BDepartamento = False
                            BSeccion = False
                            BMaquina = False
                            BSistema = True
                            BEquipo = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Sistemas")
                        ElseIf OptUbi.Item(5).Value = True Then
                            BEmpresa = False
                            BDepartamento = False
                            BSeccion = False
                            BMaquina = False
                            BSistema = False
                            BEquipo = True
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From M_Equipos")
                        End If
                                
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "4000"
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
                        
            End If

End Sub

Private Sub List1_DblClick()
          'DbGrid.Columns(4).Text = Mid(List1.Text, 1, 9)
          DbGrid.Columns(2).Text = List1.Text
          List1.Visible = False
          DbGrid.SetFocus

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
      If KeyAscii = 43 Then
          DbGrid.Columns(2).Text = List1.Text
          List1.Visible = False
          DbGrid.SetFocus
      End If
End Sub

Private Sub List1_LostFocus()
          List1.Visible = False
End Sub

