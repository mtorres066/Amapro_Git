VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Variables 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle De Catalogo"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8490
   Icon            =   "Variables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRutinas 
      Height          =   6015
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7560
         Picture         =   "Variables.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.Data DataRutinas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Rutinas"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSDBGrid.DBGrid DbGridRutinas 
         Bindings        =   "Variables.frx":0D0C
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "Variables.frx":0D26
         TabIndex        =   26
         Top             =   240
         Width           =   7215
      End
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   5175
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Variables.frx":1DA3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameVariables"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Variables.frx":20BD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Variables.frx":250F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "FrameOpciones"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar"
      Tab(2).Control(4)=   "CmdActualizar"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Buscar Todos"
         Height          =   735
         Left            =   -69600
         Picture         =   "Variables.frx":2961
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4200
         Width           =   2535
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar Datos"
         Height          =   735
         Left            =   -69600
         Picture         =   "Variables.frx":2C6B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Frame FrameVariables 
         Caption         =   "Datos De Los Estandares De Dimensionales Por Catalogo Y la Rutina"
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
         Height          =   3495
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   7995
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MaximoClientePulgadas"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   5
            Top             =   2760
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MaximoInternoPulgadas"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   4
            Top             =   2400
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MinimoInternoPulgadas"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   3
            Top             =   2040
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MinimoClientePulgadas"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   2
            Top             =   1680
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MaximoClienteMilimetros"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   7
            Left            =   5400
            TabIndex        =   9
            Top             =   2760
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MaximoInternoMilimetros"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   8
            Top             =   2400
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MinimoInternoMilimetros"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   5
            Left            =   5400
            TabIndex        =   7
            Top             =   2040
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMedidas 
            DataField       =   "MinimoClienteMilimetros"
            DataSource      =   "DataVariables"
            Height          =   285
            Index           =   4
            Left            =   5400
            TabIndex        =   6
            Top             =   1680
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Codigo"
            DataSource      =   "DataVariables"
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
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox TxtRut 
            Appearance      =   0  'Flat
            DataField       =   "Rutina"
            DataSource      =   "DataVariables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   1
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Maximo Cliente"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   43
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Maximo Interno"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   42
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Minimo Interno"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   41
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Minimo Cliente"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   40
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Medidas En Milimetros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   39
            Top             =   1320
            Width           =   2370
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Medidas En Pulgadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   600
            TabIndex        =   38
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Minimo Cliente"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   36
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Minimo Interno"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   35
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Maximo Interno"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   34
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Maximo Cliente"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   33
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label LblBarniz 
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   2880
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Rutina"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   975
         End
         Begin VB.Label LblRutina 
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
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   2760
            TabIndex        =   30
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label LblVariable 
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
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2760
            TabIndex        =   29
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68640
         TabIndex        =   20
         ToolTipText     =   " "
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones de Busqueda"
         Height          =   740
         Left            =   -74760
         TabIndex        =   17
         Top             =   1080
         Width           =   3405
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Codigo Catalogo"
            Height          =   225
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptCodigoRutina 
            Caption         =   "Codigo Rutina"
            Height          =   195
            Left            =   1920
            TabIndex        =   19
            ToolTipText     =   " "
            Top             =   360
            Width           =   1340
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Variables.frx":30AD
         Height          =   4305
         Left            =   -74880
         OleObjectBlob   =   "Variables.frx":30C9
         TabIndex        =   16
         Top             =   720
         Width           =   8025
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo Catalogo"
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
         Left            =   -70155
         TabIndex        =   28
         Top             =   2040
         Width           =   1410
      End
   End
   Begin VB.Data DataVariables 
      BackColor       =   &H80000014&
      Caption         =   "Detalle De Catalogos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Erick\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VariablesMedia"
      Top             =   6120
      Width           =   8265
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   600
      Left            =   6840
      MouseIcon       =   "Variables.frx":48CC
      Picture         =   "Variables.frx":4D0E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   1300
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   600
      Left            =   5520
      MouseIcon       =   "Variables.frx":5150
      Picture         =   "Variables.frx":5592
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1300
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   600
      Left            =   4200
      MouseIcon       =   "Variables.frx":5AC4
      Picture         =   "Variables.frx":5F06
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1300
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   600
      Left            =   2880
      MouseIcon       =   "Variables.frx":6438
      Picture         =   "Variables.frx":687A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1300
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   600
      Left            =   1560
      MouseIcon       =   "Variables.frx":6DAC
      Picture         =   "Variables.frx":71EE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1300
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   600
      Left            =   240
      MouseIcon       =   "Variables.frx":7720
      Picture         =   "Variables.frx":7B62
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1300
   End
End
Attribute VB_Name = "Variables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim RRutinas As Recordset

Dim RVariable As Recordset

Sub botones()
    If Bandera = True Then
         FrameVariables.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataVariables.Visible = False
         FrameOpciones.Visible = False
         DBGrid1.Visible = False
    Else
         FrameVariables.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataVariables.Visible = True
         FrameOpciones.Visible = True
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
        DataVariables.RecordSource = ("Select * from VariablesMedia")
        DataVariables.Refresh
        DBGrid1.Refresh
        TabDetalle.Tab = 1
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        'AGREGA UN REGISTRO
        DataVariables.Recordset.AddNew
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
            Bandera = True
            botones
            TxtCod.SetFocus
        End If
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataVariables.Recordset.Delete
                DataVariables.Recordset.MoveLast
            End If
  
            If DataVariables.Recordset.EOF Then
                DataVariables.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdBuscar_Click()
        If OptCodigo.Value = True Then
            DataVariables.RecordSource = ("Select * from VariablesMedia where Codigo like '" & TxtBuscar.Text & "*'")
        ElseIf OptCodigoRutina.Value = True Then
            DataVariables.RecordSource = ("Select * from VariablesMedia where Rutina like '" & TxtBuscar.Text & "*'")
        End If
            DataVariables.Refresh
            DBGrid1.Refresh
            TabDetalle.Tab = 1
End Sub

Private Sub CmdCancelar_Click()
    On Error Resume Next
        'CANCELA LOS CAMBIOS
        DataVariables.Recordset.CancelUpdate
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
            Bandera = False
            botones
        End If
End Sub

Private Sub CmdEditar_Click()
    On Error Resume Next
        'EDITA EL REGISTRO
        DataVariables.Recordset.Edit
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        Else
            Bandera = True
            botones
            TxtCod.SetFocus
        End If
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   'GRABA DATOS
   DataVariables.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Codigo de Variable ya existe", vbOKOnly + vbInformation, "Informacion"
      TxtCod.SetFocus
   ElseIf Err <> 3022 And Err <> 0 Then
      MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
   Else
      Bandera = False
      botones
  End If
      
   
      

End Sub

Private Sub CmdSale_Click()
FrameRutinas.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub



Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataVariables.RecordSource = ("Select * from VariablesMedia order by " & DBGrid1.Columns(ColIndex).DataField)
    DataVariables.Refresh
    DBGrid1.Refresh
    
End Sub

Private Sub DbGridRutinas_DblClick()
    TxtRut.Text = DbGridRutinas.Columns(0)
    FrameRutinas.Visible = False
    TxtRut.SetFocus
End Sub

Private Sub DbGridRutinas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        TxtRut.Text = DbGridRutinas.Columns(0)
        FrameRutinas.Visible = False
        TxtRut.SetFocus
    End If
End Sub

Private Sub Form_Load()
    DataVariables.DatabaseName = BasedeDatos
    DataRutinas.DatabaseName = BasedeDatos
End Sub


Private Sub MskMedidas_GotFocus(Index As Integer)
    MskMedidas.Item(Index).SelStart = 0
    MskMedidas.Item(Index).SelLength = Len(MskMedidas.Item(Index).Text)
End Sub

Private Sub MskMedidas_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptCodigo_Click()
    Lbletiqueta.Caption = "Codigo Catalogo"
    TxtBuscar.SetFocus
End Sub
Private Sub OptCodigoRutina_Click()
    Lbletiqueta.Caption = "Codigo Rutina"
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

Private Sub TxtCod_Change()
    Set RVariable = Db.OpenRecordset("Select DescripcionVariable From VariablesDescripcion Where CodigoVariable = '" & TxtCod.Text & "'")
    If RVariable.RecordCount > 0 Then
        If IsNull(RVariable(0)) Then
            LblVariable.Caption = ""
        Else
            LblVariable.Caption = RVariable(0)
        End If
    Else
        LblVariable.Caption = ""
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



Private Sub TxtRut_Change()
    Set RRutinas = Db.OpenRecordset("Select Descrip From Rutinas Where Rutina = '" & TxtRut.Text & "'")
    If RRutinas.RecordCount > 0 Then
        LblRutina.Caption = RRutinas(0)
    Else
        LblRutina.Caption = ""
    End If
    
End Sub

Private Sub TxtRut_DblClick()
    DataRutinas.RecordSource = ("Select * From Rutinas")
    DataRutinas.Refresh
    DbGridRutinas.Refresh
    FrameRutinas.Visible = True
    DbGridRutinas.SetFocus
End Sub

Private Sub TxtRut_GotFocus()
        TxtRut.SelStart = 0
        TxtRut.SelLength = Len(TxtRut.Text)
End Sub

Private Sub TxtRut_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
        
    If KeyAscii = 43 Then
        DataRutinas.RecordSource = ("Select * From Rutinas")
        DataRutinas.Refresh
        DbGridRutinas.Refresh
        FrameRutinas.Visible = True
        DbGridRutinas.SetFocus
    End If
End Sub
