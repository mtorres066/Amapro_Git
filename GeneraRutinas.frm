VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form GeneraRutinas 
   Caption         =   "Generar Rutinas"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "GeneraRutinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   9
      Top             =   -240
      Width           =   8655
      Begin MSMask.MaskEdBox TxtHor 
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opciones De Tipo De Rutina "
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
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   6375
         Begin VB.OptionButton Option1 
            Caption         =   "Rutinas &Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4200
            Picture         =   "GeneraRutinas.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   2040
         End
         Begin VB.OptionButton OptProceso 
            Caption         =   "Rutinas De &Proceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   120
            Picture         =   "GeneraRutinas.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Solo Rutinas Seleccionadas"
            Top             =   240
            Value           =   -1  'True
            Width           =   1920
         End
         Begin VB.OptionButton OptArranque 
            Caption         =   "Rutinas De &Arranque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   2160
            Picture         =   "GeneraRutinas.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Todas Las Rutinas"
            Top             =   240
            Width           =   1920
         End
      End
      Begin VB.TextBox TxtCatalogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MskFec 
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   2160
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtLin 
         Appearance      =   0  'Flat
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
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Top             =   2160
         Width           =   525
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6720
         Picture         =   "GeneraRutinas.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Borra Totas Rutinas Seleccionadas"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton CmdSalida 
         Cancel          =   -1  'True
         Caption         =   "&Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6720
         Picture         =   "GeneraRutinas.frx":1412
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salida"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "&Generar Rutina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6720
         Picture         =   "GeneraRutinas.frx":3484
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Genera Rutinas"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Informacion de Proceso"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   6375
         Begin VB.Label LblRutinas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4920
            TabIndex        =   16
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label LblFinal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label LblInicio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rutinas Generadas"
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
            Left            =   3240
            TabIndex        =   20
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hora Final"
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
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Hora Inicio"
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
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   4095
         Left            =   6600
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
         Caption         =   "Catalogo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   26
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label LblOrigen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label LblFicha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   2520
         Width           =   5415
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000004&
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000004&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   615
      End
   End
End
Attribute VB_Name = "GeneraRutinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaMaestroRutinas As New ADODB.Recordset
Dim RBuscaRutinas As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaRutinasDeCatalogo As New ADODB.Recordset
Dim RBuscaCatalogo As New ADODB.Recordset
Dim VFechaActual As Date
Dim VHoraActual As String
Dim VCatalogo As String
Dim Contadorcabezales As Double
Dim Cont As Double
Dim ContadorRutinas As Double
Dim RCuentaRutinas As New ADODB.Recordset
Dim RCuentaLineas As New ADODB.Recordset
Dim RCuentaCabezales As New ADODB.Recordset
Dim mensaje As String

Private Sub CmdBorrar_Click()
On Error Resume Next
        If GOrigenDeDatos = "AmaproAccess" Then
        Else
            MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
        End If
        
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    mensaje = MsgBox("¿Está seguro de Borrar Estas Rutinas?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
    If mensaje = 1 Then
        If GOrigenDeDatos = "AmaproAccess" Then
            Conexion.Execute ("delete From CapturaRutinas Where Linea = '" & TxtLin.Text & "' and Fec_rut = # " & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_rut = '" & TxtHor.Text & "'")
        Else
            Conexion.Execute ("delete From CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and Fec_rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_rut) = '" & UCase(TxtHor.Text) & "'")
        End If
        If Err <> 0 Then
            MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
        MsgBox "Rutinas Borradas", vbOKOnly + vbInformation, "Informacion"
    End If
    
End Sub

Private Sub CmdGenerar_Click()
On Error Resume Next
MousePointer = 11

    
    LblInicio.Caption = Time
    
    
    'SI LA HORA ES MENOR QUE LAS 7 DE LA MAÑANA ENTONCES DA LA FECHA ANTERIOR
     If Format(Time, "hh:mm") < "07:00" Then
        VFechaActual = Format(DateValue(MskFec.Text) - 1, "dd/mm/yyyy")
     Else
        VFechaActual = Format(MskFec.Text, "dd/mm/yyyy")
     End If
     

    'HORA
     VHoraActual = Format(TxtHor.Text, "hh:mm")
    'CATALOGO DE FICHA TECNICA
     VCatalogo = TxtCatalogo.Text
   
        
    'BUSCA LINEA
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Linea, Esp_Tec From Lineas Where Linea = '" & TxtLin.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Linea, Esp_Tec From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
    
        If RBuscaLinea.RecordCount > 0 Then
        Else
            MsgBox "Codigo De Linea No Existe", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
        
        
        
        
    'VERIFICA SI YA EXISTEN RUTINAS
    Set RBuscaRutinas = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaRutinas, "Select * From CapturaRutinas Where Linea = '" & TxtLin.Text & "' and Fec_rut = # " & Format(MskFec.Text, "mm/dd/yyyy") & "# and Hor_rut = '" & TxtHor.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaRutinas, "Select * From CapturaRutinas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and Fec_rut = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & " and UPPER(Hor_rut) = '" & UCase(TxtHor.Text) & "'")
        End If
        If RBuscaRutinas.RecordCount > 0 Then
                MsgBox "Rutinas Ya Existen Verifique", vbOKOnly + vbCritical, "Verifique"
                TxtLin.SetFocus
                MousePointer = 0
                Exit Sub
        End If
        
        
        

        ContadorRutinas = 0
        
        'SELECCIONA TODAS LAS RUTINAS QUE TIENE EL CATALOGO DEPENDIENDE DE LA FICHA TECNICA Y SI
        'SON RUTINAS DE PROCESO
        Set RBuscaRutinasDeCatalogo = New ADODB.Recordset
        
        If OptProceso.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM, Rutinas R Where VM.Codigo = '" & VCatalogo & "' And VM.Rutina = R.Rutina And R.GeneraRutinaProceso = -1")
                Else
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM, Rutinas R Where UPPER(VM.Codigo) = '" & UCase(VCatalogo) & "' And UPPER(VM.Rutina) = UPPER(R.Rutina) And R.GeneraRutinaProceso = -1")
                End If
        'SON RUTINAS DE ARRANQUE
        ElseIf OptArranque.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM, Rutinas R Where VM.Codigo = '" & VCatalogo & "' And VM.Rutina = R.Rutina And R.GeneraRutinaArranque = -1")
                Else
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM, Rutinas R Where UPPER(VM.Codigo) = '" & UCase(VCatalogo) & "' And UPPER(VM.Rutina) = UPPER(R.Rutina) And R.GeneraRutinaArranque = -1")
                End If
                    
        'TODAS LAS RUTINAS DEL CATALOGO
        Else
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM Where VM.Codigo = '" & VCatalogo & "'")
                Else
                    Call Abrir_Recordset(RBuscaRutinasDeCatalogo, "Select VM.Rutina, VM.Cabezales From VariablesMedia VM Where UPPER(VM.Codigo) = '" & UCase(VCatalogo) & "'")
                End If
        End If
        
        
                        
            'SI ENCUENTRA RUTINAS PARA ESTA LINEA
            If RBuscaRutinasDeCatalogo.RecordCount > 0 Then
            
                        
                        'CREA UN CICLO POR CADA RUTINA QUE TENGA LA LINEA ASIGINADA
                        Do Until RBuscaRutinasDeCatalogo.EOF
                            Cont = 1
                            'BUSCA CUANTOS CABEZALES TIENE ASIGNADA LA RUTINA
'                            Set RBuscaMaestroRutinas = New ADODB.Recordset
'                                If GOrigenDeDatos = "AmaproAccess" Then
'                                    Call Abrir_Recordset(RBuscaMaestroRutinas, "Select Rutina, Cabezal From Rutinas Where Rutina = '" & RBuscaRutinasDeCatalogo!Rutina & "'")
'                                Else
'                                    Call Abrir_Recordset(RBuscaMaestroRutinas, "Select Rutina, Cabezal From Rutinas Where UPPER(Rutina) = '" & UCase(RBuscaRutinasDeCatalogo!Rutina) & "'")
'                                End If
'                                If RBuscaMaestroRutinas.RecordCount > 0 Then
                                    'CANTIDAD DE CABEZALES QUE TIENE LA RUTINA EN FICHA TECNICA DE RUTINAS
                                    'Contadorcabezales = RBuscaMaestroRutinas!cabezal
                                    Contadorcabezales = RBuscaRutinasDeCatalogo!cabezales
                                    'INICIA LA TRANSACCION
                                    Conexion.BeginTrans
                                    Do While Cont <= Contadorcabezales
                                                         'AGREGA DATOS A LA BASE DE CAPTURA DE DATOS EN BASE AL NUMERO DE CABEZALES
                                                          If GOrigenDeDatos = "AmaproAccess" Then
                                                                Conexion.Execute "Insert Into CapturaRutinas (Linea, Fec_Rut, Hor_Rut, Esp_Tec, Cabezal, Rutina, Valor, Catalogo) VALUES ('" & RBuscaLinea!Linea & "', #" & Format(VFechaActual, "mm/dd/yyyy") & "#, '" & VHoraActual & "', '" & RBuscaLinea!Esp_Tec & "', " & Cont & ", '" & RBuscaRutinasDeCatalogo!Rutina & "', 0, '" & VCatalogo & "')"
                                                          Else
                                                                Conexion.Execute "Insert Into CapturaRutinas (Linea, Fec_Rut, Hor_Rut, Esp_Tec, Cabezal, Rutina, Valor, Catalogo) VALUES ('" & RBuscaLinea!Linea & "', To_Date('" & VFechaActual & "', 'dd/mm/yyyy')" & ", '" & VHoraActual & "', '" & RBuscaLinea!Esp_Tec & "', " & Cont & ", '" & RBuscaRutinasDeCatalogo!Rutina & "', 0, '" & VCatalogo & "')"
                                                          End If
                                                          If Err <> 0 Then
                                                             Conexion.RollbackTrans
                                                             MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation + vbCritical & "Error"
                                                             Err.Clear
                                                          End If
                                                          
                                                      Cont = Cont + 1
                                                      ContadorRutinas = ContadorRutinas + 1
                                    Loop
                                    'TERMINA LA TRANSACCION
                                    Conexion.CommitTrans
                                    
                                'End If
                                RBuscaRutinasDeCatalogo.MoveNext
                        Loop
            Else
                    MsgBox "El Catalogo Que Tiene Esta Ficha Tecnica No Tiene Asignadas Rutinas", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
            End If
                        
              
LblFinal.Caption = Time
LblRutinas.Caption = ContadorRutinas

MousePointer = 0

MsgBox "Rutinas Generadas Con Exito", vbOKOnly + vbInformation, "Informacion"

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    TxtLin.SetFocus
End Sub

Private Sub Form_Load()
     MskFec.Text = Date
     TxtHor.Text = Format(Time, "hh:mm")

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

Private Sub TxtHor_GotFocus()
        TxtHor.SelStart = 0
        TxtHor.SelLength = Len(TxtHor.Text)
End Sub

Private Sub TxtHor_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_Change()
    'SELECCIONA LA FICHA TECNICA DE LA LINEA
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where Linea = '" & TxtLin.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
        End If
    If RBuscaLinea.RecordCount > 0 Then
        'BUSCA LA DESCRIPCION Y ORIGEN DE FICHA TECNICA
        Set RBuscaFicha = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select * from FichaTecnica Where Esp_Tec = '" & RBuscaLinea!Esp_Tec & "'")
            Else
                    Call Abrir_Recordset(RBuscaFicha, "Select * from FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(RBuscaLinea!Esp_Tec) & "'")
            End If
            If RBuscaFicha.RecordCount > 0 Then
                LblFicha.Caption = RBuscaFicha!Esp_Tec & "/" & RBuscaFicha!Descrip
                If IsNull(RBuscaFicha!Origen) Then
                    LblOrigen.Caption = ""
                    TxtCatalogo.Text = ""
                Else
                    LblOrigen.Caption = RBuscaFicha!Origen
                    TxtCatalogo.Text = RBuscaFicha!Variables
                End If
            Else
                LblFicha.Caption = ""
                LblOrigen.Caption = ""
                TxtCatalogo.Text = ""
            End If
    Else
                LblFicha.Caption = ""
                LblOrigen.Caption = ""
                TxtCatalogo.Text = ""
    End If
        
End Sub

Private Sub TxtLin_GotFocus()
    TxtLin.SelStart = 0
    TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
