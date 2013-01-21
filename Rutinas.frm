VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Rutinas 
   BackColor       =   &H000000FF&
   Caption         =   "Rutinas"
   ClientHeight    =   7995
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10545
   Icon            =   "Rutinas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabRutinas 
      Height          =   6855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Rutinas.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameRutinas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DbGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSDataGridLib.DataGrid DbGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            DataField       =   "Rutina"
            Caption         =   "Rutina"
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
            DataField       =   "Descrip"
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
            DataField       =   "Cabezal"
            Caption         =   "Cabezal"
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
            DataField       =   "Imp_Rut"
            Caption         =   "Imp_Rut"
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
            DataField       =   "GeneraRutinaProceso"
            Caption         =   "Rutina Proceso"
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
            DataField       =   "GeneraRutinaArranque"
            Caption         =   "Rutina Arranque"
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
            DataField       =   "Grupo"
            Caption         =   "Grupo"
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
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3974.74
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameRutinas 
         Caption         =   "Datos Generales De La Rutina"
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
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   9975
         Begin VB.TextBox TxtGru 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Esta Rutina Se Incluye En Generacion De Rutinas De Arranque"
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
            Left            =   1800
            TabIndex        =   6
            Top             =   2160
            Width           =   6375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Esta Rutina Se Incluye En Generacion De Rutinas De Proceso"
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
            Left            =   1800
            TabIndex        =   5
            Top             =   1800
            Width           =   6255
         End
         Begin VB.TextBox TxtRut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   7935
         End
         Begin VB.TextBox TxtCab 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7800
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox ChkImpRut 
            Caption         =   "Imprime Rutina En Reporte De Batch"
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
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   435
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   21
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Rutina"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cabezales De Rutina"
            Height          =   195
            Index           =   20
            Left            =   6240
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label LblTapa 
            Height          =   195
            Left            =   3000
            TabIndex        =   16
            Top             =   3120
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   840
      Left            =   8640
      MouseIcon       =   "Rutinas.frx":045E
      Picture         =   "Rutinas.frx":08A0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1725
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   840
      Left            =   6960
      MouseIcon       =   "Rutinas.frx":0DBB
      Picture         =   "Rutinas.frx":11FD
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   840
      Left            =   5280
      MouseIcon       =   "Rutinas.frx":17C5
      Picture         =   "Rutinas.frx":1C07
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   840
      Left            =   3600
      MouseIcon       =   "Rutinas.frx":213E
      Picture         =   "Rutinas.frx":2580
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   840
      Left            =   1920
      MouseIcon       =   "Rutinas.frx":2ADC
      Picture         =   "Rutinas.frx":2F1E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   840
      Left            =   240
      MouseIcon       =   "Rutinas.frx":32F5
      Picture         =   "Rutinas.frx":3737
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1600
   End
End
Attribute VB_Name = "Rutinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BEditar As Boolean
Dim RRutinas As New ADODB.Recordset
Dim vtexto As String


Sub botones()
    If Bandera = True Then
         FrameRutinas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtRut.SetFocus

         DbGrid1.Visible = False
    Else
         FrameRutinas.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True

         DbGrid1.Visible = True
    End If
End Sub




Private Sub Check1_KeyPress(KeyAscii As Integer)
             If KeyAscii = 13 Then
                 SendKeys "{tab}"
             End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If

End Sub

Private Sub ChkImpRut_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If

End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
            Bandera = True
            botones
            Limpia_Campos
            TxtRut.SetFocus
            TxtUsu.Text = GUsuario
            BEditar = False
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            If (TxtRut.Text = "0500" Or TxtRut.Text = "0501" Or TxtRut.Text = "0600" Or TxtRut.Text = "0800" Or TxtRut.Text = "0900" Or TxtRut.Text = "1001" Or TxtRut.Text = "1002" Or TxtRut.Text = "3001" Or TxtRut.Text = "3002") Then
                MsgBox "Esta Rutina " & TxtRut.Text & " No Se Puede Borrar Porque La Utiliza El Software SEAMETAL O CPA, Consulte al Programador", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            
            If (TxtRut.Text = "1000" Or TxtRut.Text = "2000" Or TxtRut.Text = "3000" Or TxtRut.Text = "4000" Or TxtRut.Text = "5000" Or TxtRut.Text = "6000" Or TxtRut.Text = "7000" Or TxtRut.Text = "7250" Or TxtRut.Text = "7500") Then
                MsgBox "Esta Rutina " & TxtRut.Text & " No Se Puede Borrar Porque La Utiliza El Software SEAMETAL O CPA, Consulte al Programador", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                    'BORRA EL REGISTRO
                        RRutinas.Delete
                        
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
                        RRutinas.Requery
                        Set DbGrid1.DataSource = RRutinas
                        'MUEVE AL SIGUIENTE REGISTRO
                        RRutinas.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
            End If
            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next
            Bandera = False
            botones
            Llena_Campos
        
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
            Bandera = True
            botones
            TxtRut.Enabled = False
            TxtDes.SetFocus
            TxtUsu.Text = GUsuario
            BEditar = True
        
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
                If Not IsNumeric(TxtCab.Text) Then
                        MsgBox "Cabezales Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
   
                'AGREGAR
                    If BEditar = False Then
                            vtexto = "Values('" & TxtRut.Text & "', '" ' CODIGO
                            vtexto = vtexto & TxtDes.Text & "', " 'DESCRIPCION
                            vtexto = vtexto & "0" & ", " 'CABEZAL
                            If ChkImpRut.Value = "1" Then
                                vtexto = vtexto & "-1" & ", " 'IMPRIME RUTINA
                            Else
                                vtexto = vtexto & "0" & ", " '
                            End If
                            If Check1.Value = "1" Then
                                vtexto = vtexto & "-1" & ", " 'ES RUTINA PROCESO
                            Else
                                vtexto = vtexto & "0" & ", " '
                            End If
                            If Check2.Value = "1" Then
                                vtexto = vtexto & "-1" & ", '" 'ES RUTINA DE ARRANGQUE
                            Else
                                vtexto = vtexto & "0" & ", '"
                            End If
                            vtexto = vtexto & TxtUsu.Text & "', '" 'USUARIO
                            vtexto = vtexto & TxtGru.Text & "')" 'GRUPO
                            
                            Conexion.Execute "Insert Into Rutinas " & vtexto
                    'EDITAR
                    Else
                            vtexto = "Descrip = '" & TxtDes.Text & "', " 'DESCRIPCION
                            vtexto = vtexto & "Cabezal = " & "0" & ", " 'CABEZAL
                            If ChkImpRut.Value = "1" Then
                                vtexto = vtexto & "Imp_Rut = -1" & ", " 'IMPRIME RUTINA
                            Else
                                vtexto = vtexto & "Imp_Rut = 0" & ", "
                            End If
                            If Check1.Value = "1" Then
                                vtexto = vtexto & "GeneraRutinaProceso = -1" & ", " 'RUTINA PROCESO
                            Else
                                vtexto = vtexto & "GeneraRutinaProceso = 0" & ", "
                            End If
                            If Check2.Value = "1" Then
                                vtexto = vtexto & "GeneraRutinaArranque = -1" & ", " 'RUTINA ARRANQUE
                            Else
                                vtexto = vtexto & "GeneraRutinaArranque = 0" & ", "
                            End If
                            vtexto = vtexto & "usuario = '" & TxtUsu.Text & "', " ' USUARIO
                            vtexto = vtexto & "Grupo = '" & TxtGru.Text & "'" ' GRUPO
                            vtexto = vtexto & "Where Rutina = '" & TxtRut.Text & "'"
                        
                            Conexion.Execute "UPDATE Rutinas SET " & vtexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtRut.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtRut.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                    
   
   
   
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                        TxtRut.Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RRutinas.Requery
                        RRutinas.MoveLast
                        Llena_Campos
   
   
        
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
        RRutinas.Sort = RRutinas.Fields(ColIndex).Name
End Sub


Private Sub DbGrid1_SelChange(Cancel As Integer)
                Llena_Campos
End Sub

Private Sub Form_Load()
        Set RRutinas = New ADODB.Recordset
        Call Abrir_Recordset(RRutinas, "Select * From Rutinas")
        Set DbGrid1.DataSource = RRutinas
        Llena_Campos
    
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

Private Sub TxtCab_GotFocus()
        TxtCab.SelStart = 0
        TxtCab.SelLength = Len(TxtCab.Text)
End Sub

Private Sub TxtCab_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtGru_GotFocus()
        TxtGru.SelStart = 0
        TxtGru.SelLength = Len(TxtGru.Text)
End Sub

Private Sub TxtGru_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtRut_GotFocus()
        TxtRut.SelStart = 0
        TxtRut.SelLength = Len(TxtRut.Text)
End Sub

Private Sub TxtRut_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Public Sub Limpia_Campos()
        TxtRut.Text = ""
        TxtDes.Text = ""
        TxtCab.Text = 0
        ChkImpRut.Value = 0
        Check1.Value = 0
        Check2.Value = 0
        TxtUsu.Text = ""
        TxtGru.Text = ""
        
End Sub


Public Sub Llena_Campos()
On Error Resume Next
        
            TxtRut.Text = RRutinas!Rutina
        
            If IsNull(RRutinas!Descrip) Then
                TxtDes.Text = ""
            Else
                TxtDes.Text = RRutinas!Descrip
            End If
            TxtCab.Text = RRutinas!cabezal
            
        
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If RRutinas!Imp_Rut = "Verdadero" Then
                                ChkImpRut.Value = "1"
                            Else
                                ChkImpRut.Value = "0"
                            End If
                    Else
                            If RRutinas!Imp_Rut = "-1" Then
                                ChkImpRut.Value = "1"
                            Else
                                ChkImpRut.Value = "0"
                            End If
                    End If
                            
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If RRutinas!GeneraRutinaProceso = "Verdadero" Then
                                Check1.Value = "1"
                            Else
                                Check1.Value = "0"
                            End If
                    Else
                            If RRutinas!GeneraRutinaProceso = "-1" Then
                                Check1.Value = "1"
                            Else
                                Check1.Value = "0"
                            End If
                    End If
                    
                    If GOrigenDeDatos = "AmaproAccess" Then
                            If RRutinas!GeneraRutinaArranque = "Verdadero" Then
                                Check2.Value = "1"
                            Else
                                Check2.Value = "0"
                            End If
                    Else
                            If RRutinas!GeneraRutinaArranque = "-1" Then
                                Check2.Value = "1"
                            Else
                                Check2.Value = "0"
                            End If
                    End If
                    
                    TxtUsu.Text = RRutinas!Usuario
                    If IsNull(RRutinas!Grupo) Then
                        TxtGru.Text = ""
                    Else
                        TxtGru.Text = RRutinas!Grupo
                    End If
                    
                    If Err <> 0 Then
                    End If

End Sub
