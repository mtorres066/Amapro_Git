VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpleadosHorasAutomatico 
   BackColor       =   &H0080C0FF&
   Caption         =   "Generar  Horas Automatico"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   Icon            =   "EmpleadosHorasAutomatico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Framebuscar 
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
      Height          =   5055
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7680
         Picture         =   "EmpleadosHorasAutomatico.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
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
         Height          =   735
         Left            =   3840
         TabIndex        =   34
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   35
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "EmpleadosHorasAutomatico.frx":237C
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "EmpleadosHorasAutomatico.frx":2395
         TabIndex        =   41
         Top             =   1080
         Width           =   8415
      End
      Begin VB.Label LblBusqueda 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   735
      Left            =   7800
      Picture         =   "EmpleadosHorasAutomatico.frx":2D6D
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Generar"
      Height          =   735
      Left            =   6720
      Picture         =   "EmpleadosHorasAutomatico.frx":4DDF
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opciones "
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H0080C0FF&
         Caption         =   "Empleado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H0080C0FF&
         Caption         =   "Departamento"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H0080C0FF&
         Caption         =   "Equipo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61669379
      CurrentDate     =   37873
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61669379
      CurrentDate     =   37873
   End
   Begin VB.TextBox TxtHorExtDobRea 
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Text            =   "0"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox TxtHorExtDobPro 
      Height          =   285
      Left            =   7440
      TabIndex        =   14
      Text            =   "0"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtHorExtReaNoc 
      Height          =   285
      Left            =   7440
      TabIndex        =   13
      Text            =   "0"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtHorExtProNoc 
      Height          =   285
      Left            =   7440
      TabIndex        =   12
      Text            =   "0"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtHorExtReaDiu 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Text            =   "0"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox TxtHorExtProDiu 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Text            =   "0"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtHorLabNoc 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtHorLabDiu 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Text            =   "0"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtLin 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtTur 
      Height          =   285
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label LblLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   4560
      TabIndex        =   32
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label LblDes 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label LblEti 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Empleado"
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
      Left            =   1800
      TabIndex        =   30
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Hasta Fecha"
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   29
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Turno"
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   28
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Linea"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   27
      Top             =   2160
      Width           =   390
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Laboradas Diurnas"
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   26
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Laboradas Nocturnas"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   25
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Hor. Ext. Proyectadas Diurnas"
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
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   24
      Top             =   3480
      Width           =   2580
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Extras Reales Diurnas"
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
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   23
      Top             =   3840
      Width           =   2445
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Hor. Ext. Proyectadas Nocturnas"
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
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   22
      Top             =   2640
      Width           =   2805
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Extras Reales Nocturnas"
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
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   21
      Top             =   3000
      Width           =   2670
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Extras Dobles Proyectadas"
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
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   20
      Top             =   3480
      Width           =   2850
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Horas Extras Dobles Reales"
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
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   3840
      Width           =   2385
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Desde Fecha"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   1080
      Width           =   960
   End
End
Attribute VB_Name = "EmpleadosHorasAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VFechaInicial As Date
Dim VFechaFinal As Date

Dim RBuscaLinea As Recordset
Dim RBusca As Recordset
Dim RBuscaEmpleado As Recordset
Dim RCapturaHoras As Recordset
Dim RBuscaSueldoBase As Recordset
Dim RBuscaFactores As Recordset

Dim VDiasMes As Integer

Dim VValorHoraLaboradaDiurna As Currency
Dim VValorHoraLaboradaNocturna As Currency
Dim VValorHoraExtraDiurna As Currency
Dim VValorHoraExtraNocturna As Currency
Dim VSueldoBase As Currency
Dim VFHorasDiurnas As Single
Dim VFHorasNocturnas As Single
Dim VFPorcentajeDiurnas As Single
Dim VFPorcentajeNocturnas As Single

Dim VMonHorLabDiu As Currency
Dim VMonHorLabNoc As Currency
Dim VMonHorExtDiu As Currency
Dim VMonHorExtNoc As Currency
Dim VMonHorExtDob As Currency

Dim BEquipo As Boolean
Dim BDepartamento As Boolean
Dim BLinea As Boolean
Dim BEmpleado As Boolean


Private Sub CmdGenerar_Click()
On Error Resume Next

MousePointer = 11

        VFechaInicial = DTPFecIni.Value
        VFechaFinal = DTPFecFin.Value

        'INICIALIZA UN RECORDSET PARA LUEGO AGREGAR DATOS
        Set RCapturaHoras = Db.OpenRecordset("Select * From EmpleadosCapturaHoras")
        
        'EQUIPO
        If OptOpcion.Item(0).Value = True Then
                Set RBuscaEmpleado = Db.OpenRecordset("Select * From Empleados Where Grupo = '" & TxtTexto.Text & "'")
        'DEPARTAMENTO
        ElseIf OptOpcion.Item(1).Value = True Then
                Set RBuscaEmpleado = Db.OpenRecordset("Select * From Empleados Where Departamento = '" & TxtTexto.Text & "'")
        'EMPLEADO
        ElseIf OptOpcion.Item(2).Value = True Then
                Set RBuscaEmpleado = Db.OpenRecordset("Select * From Empleados Where Codigo = '" & TxtTexto.Text & "'")
            
        End If
        
        'SI ENCUENTRA EMPLEADOS
        If RBuscaEmpleado.RecordCount > 0 Then
        
        Else
            If OptOpcion.Item(0).Value = True Then
                MsgBox "No Hay Empleados En Este Equipo", vbOKOnly + vbInformation, "Informacion"
            ElseIf OptOpcion.Item(1).Value = True Then
                MsgBox "No Hay Empleados En Este Departamento", vbOKOnly + vbInformation, "Informacion"
            ElseIf OptOpcion.Item(2).Value = True Then
                MsgBox "Empleado No Existe", vbOKOnly + vbInformation, "Informacion"
            End If
            MousePointer = 0
            Exit Sub
        End If
                
                
        'BUSCA LOS FACTORES PARA CALCULOS DE HORAS
        Set RBuscaFactores = Db.OpenRecordset("Select * From EmpleadosFactores")
            If RBuscaFactores.RecordCount > 0 Then
               VFHorasDiurnas = RBuscaFactores(0)
               VFHorasNocturnas = RBuscaFactores(1)
               VFPorcentajeDiurnas = RBuscaFactores(2)
               VFPorcentajeNocturnas = RBuscaFactores(3)
            End If
                
        'CREA UN CICLO CON EL RANGO DE FECHAS
        Do Until VFechaInicial > VFechaFinal
        
                    'MUEVE AL PRIMER REGISTRO DE EMPLEADOS
                    RBuscaEmpleado.MoveFirst
        
                    'BUSCA CUANTOS DIAS TIENE EL MES
                    VDiasMes = UltimoDiaMes(VFechaInicial)
                    
                            Do Until RBuscaEmpleado.EOF
                                    'BUSCA EL SUELDO BASE DEL EMPLEADO
                                    Set RBuscaSueldoBase = Db.OpenRecordset("Select SueldoBase From Empleados Where Codigo = '" & RBuscaEmpleado!Codigo & "'")
                                    If RBuscaSueldoBase.RecordCount > 0 Then
                                       VSueldoBase = RBuscaSueldoBase!SueldoBase
                                       VValorHoraLaboradaDiurna = ((VSueldoBase / VDiasMes) / VFHorasDiurnas)
                                       VValorHoraLaboradaNocturna = ((VSueldoBase / VDiasMes) / VFHorasNocturnas)
                                       VValorHoraExtraDiurna = VValorHoraLaboradaDiurna * VFPorcentajeDiurnas
                                       VValorHoraExtraNocturna = VValorHoraLaboradaNocturna * VFPorcentajeNocturnas
                                    Else
                                       VSueldoBase = "0"
                                       VValorHoraLaboradaDiurna = "0"
                                       VValorHoraLaboradaNocturna = "0"
                                       VValorHoraExtraDiurna = "0"
                                       VValorHoraExtraNocturna = "0"
                                     End If
                                                
                                                
                                       'MONTO LABORADAS DIURNAS
                                        VMonHorLabDiu = Format(TxtHorLabDiu.Text * VValorHoraLaboradaDiurna, "#,###,##0.00")
                                        'MONTO LABORADAS NOCTURNAS
                                        VMonHorLabNoc = Format(TxtHorLabNoc.Text * VValorHoraLaboradaNocturna, "#,###,##0.00")
                                        'MONTO HORAS EXTRAS DIURNAS
                                        VMonHorExtDiu = Format(TxtHorExtReaDiu.Text * VValorHoraExtraDiurna, "#,###,##0.00")
                                        'MONTO HORAS EXTRAS NOCTURNAS
                                        VMonHorExtNoc = Format(TxtHorExtReaNoc.Text * VValorHoraExtraNocturna, "#,###,##0.00")
                                        'MONTO HORAS EXTRAS DOBLES
                                        VMonHorExtDob = Format((TxtHorExtDobRea.Text * (VValorHoraExtraDiurna * 2)), "#,###,##0.00")
                                        
                                        
                                        
                                                    'AGREGA DATOS A LA BASE DE DATOS
                                                    RCapturaHoras.AddNew
                                                            RCapturaHoras!fecha = VFechaInicial
                                                            RCapturaHoras!Turno = TxtTur.Text
                                                            RCapturaHoras!Linea = TxtLin.Text
                                                            RCapturaHoras!Empleado = RBuscaEmpleado!Codigo
                                                            
                                                            RCapturaHoras!HorasLaboradasDiurnas = TxtHorLabDiu.Text
                                                            RCapturaHoras!MontoHorasLaboradasDiurnas = VMonHorLabDiu
                                                            
                                                            RCapturaHoras!HorasLaboradasNocturnas = TxtHorLabNoc.Text
                                                            RCapturaHoras!MontoHorasLaboradasNocturnas = VMonHorLabNoc
                                                            
                                                            RCapturaHoras!HorasExtrasDiurnasProyectadas = TxtHorExtProDiu.Text
                                                            RCapturaHoras!HorasExtrasDiurnas = TxtHorExtReaDiu.Text
                                                            RCapturaHoras!MontoHorasExtrasDiurnas = VMonHorExtDiu
                                                            
                                                            RCapturaHoras!HorasExtrasNocturnasProyectadas = TxtHorExtProNoc.Text
                                                            RCapturaHoras!HorasExtrasNocturnas = TxtHorExtReaNoc.Text
                                                            RCapturaHoras!MontoHorasExtrasnocturnas = VMonHorExtNoc
                                                            
                                                            RCapturaHoras!HorasExtrasDoblesProyectadas = TxtHorExtDobPro.Text
                                                            RCapturaHoras!HorasExtrasDobles = TxtHorExtDobRea.Text
                                                            RCapturaHoras!MontoHorasExtrasDobles = VMonHorExtDob
                                                            
                                                            RCapturaHoras!usuario = GUsuario
                                                    RCapturaHoras.Update
                                                    
                                                    If Err.Number <> 0 Then
                                                        MsgBox "Ya Hay Datos En Esta Fecha y Linea Y Turno Y Empleado", vbOKOnly + vbInformation, "Informacion"
                                                    End If
                    
                                RBuscaEmpleado.MoveNext
                            Loop
                
            VFechaInicial = DateValue(VFechaInicial) + 1
        Loop
        
        
MousePointer = 0

        MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
        
End Sub

Private Sub CmdSale_Click()
    Framebuscar.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridBuscar_DblClick()
  If BEquipo = True Or BDepartamento = True Or BEmpleado = True Then
            TxtTexto.Text = DBGridBuscar.Columns(0)
            TxtTexto.SetFocus
        ElseIf BLinea = True Then
            TxtLin.Text = DBGridBuscar.Columns(0)
            TxtLin.SetFocus
        End If
        Framebuscar.Visible = False
      
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BEquipo = True Or BDepartamento = True Or BEmpleado = True Then
                    TxtTexto.Text = DBGridBuscar.Columns(0)
                    TxtTexto.SetFocus
                ElseIf BLinea = True Then
                    TxtLin.Text = DBGridBuscar.Columns(0)
                    TxtLin.SetFocus
                End If
                Framebuscar.Visible = False
        End If
      
End Sub

Private Sub DtpFecFin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub DtpFecIni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Form_Load()
                'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        DataBuscar.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        DataBuscar.DatabaseName = BasedeDatos
        
        DTPFecIni.Value = Date
        DTPFecFin.Value = Date

End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            LblEti.Caption = "Equipo"
        ElseIf Index = 1 Then
            LblEti.Caption = "Departamento"
        ElseIf Index = 2 Then
            LblEti.Caption = "Empleado"
        End If
End Sub


Private Sub Txtbuscar_Change()
            
            
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                            'OPCION CUALQUIER PALABRA
                            If OptBusqueda.Item(3).Value = True Then
                                    If BEquipo = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosGrupos Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BDepartamento = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosDepartamentos Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip from Lineas Where Descrip Like '*" & TxtBuscar.Text & "*'")
                                    ElseIf BEmpleado = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from Empleados Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                                    End If
                            'OPCION PALABRA INICIAL
                            ElseIf OptBusqueda.Item(2).Value = True Then
                                    If BEquipo = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosGrupos Where Descripcion Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BDepartamento = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosDepartamentos Where Descripcion Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BLinea = True Then
                                        DataBuscar.RecordSource = ("Select Linea, Descrip from Lineas Where Descrip Like '" & TxtBuscar.Text & "*'")
                                    ElseIf BEmpleado = True Then
                                        DataBuscar.RecordSource = ("Select Codigo, Descripcion from Empleados Where Descripcion Like '" & TxtBuscar.Text & "*'")
                                    End If
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION CUALQUIER PALABRA
                            If OptBusqueda.Item(3).Value = True Then
                                If BEquipo = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosGrupos Where Codigo Like '*" & TxtBuscar.Text & "*'")
                                ElseIf BDepartamento = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosDepartamentos Where Codigo Like '*" & TxtBuscar.Text & "*'")
                                ElseIf BLinea = True Then
                                    DataBuscar.RecordSource = ("Select Linea, Descrip from Lineas Where Linea Like '*" & TxtBuscar.Text & "*'")
                                ElseIf BEmpleado = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from Empleados Where Codigo Like '*" & TxtBuscar.Text & "*'")
                                End If
                            'OPCION PALABRA INICIAL
                            ElseIf OptBusqueda.Item(2).Value = True Then
                                If BEquipo = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosGrupos Where Codigo Like '" & TxtBuscar.Text & "*'")
                                ElseIf BDepartamento = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from EmpleadosGrupos Where Codigo Like '" & TxtBuscar.Text & "*'")
                                ElseIf BLinea = True Then
                                    DataBuscar.RecordSource = ("Select Linea, Descrip from Lineas Where Linea Like '" & TxtBuscar.Text & "*'")
                                ElseIf BEmpleado = True Then
                                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from Empleados Where Codigo Like '" & TxtBuscar.Text & "*'")
                                End If
                            End If
                    End If
                            DataBuscar.Refresh
                            DBGridBuscar.Refresh
                            DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtHorExtDobPro_GotFocus()
        TxtHorExtDobPro.SelStart = 0
        TxtHorExtDobPro.SelLength = Len(TxtHorExtDobPro.Text)
End Sub

Private Sub TxtHorExtDobPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorExtDobRea_GotFocus()
        TxtHorExtDobRea.SelStart = 0
        TxtHorExtDobRea.SelLength = Len(TxtHorExtDobRea.Text)
End Sub

Private Sub TxtHorExtDobRea_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorExtProDiu_GotFocus()
        TxtHorExtProDiu.SelStart = 0
        TxtHorExtProDiu.SelLength = Len(TxtHorExtProDiu.Text)
End Sub

Private Sub TxtHorExtProDiu_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorExtProNoc_GotFocus()
        TxtHorExtProNoc.SelStart = 0
        TxtHorExtProNoc.SelLength = Len(TxtHorExtProNoc.Text)
End Sub

Private Sub TxtHorExtProNoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorExtReaDiu_GotFocus()
        TxtHorExtReaDiu.SelStart = 0
        TxtHorExtReaDiu.SelLength = Len(TxtHorExtReaDiu.Text)
End Sub

Private Sub TxtHorExtReaDiu_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorExtReaNoc_GotFocus()
        TxtHorExtReaNoc.SelStart = 0
        TxtHorExtReaNoc.SelLength = Len(TxtHorExtReaNoc.Text)
End Sub

Private Sub TxtHorExtReaNoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorLabDiu_GotFocus()
        TxtHorLabDiu.SelStart = 0
        TxtHorLabDiu.SelLength = Len(TxtHorLabDiu.Text)
End Sub

Private Sub TxtHorLabDiu_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtHorLabNoc_GotFocus()
        TxtHorLabNoc.SelStart = 0
        TxtHorLabNoc.SelLength = Len(TxtHorLabNoc.Text)
End Sub

Private Sub TxtHorLabNoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
End Sub

Private Sub TxtLin_DblClick()
        BLinea = True
        BEquipo = False
        BDepartamento = False
        DataBuscar.RecordSource = "Select Linea, Descrip From Lineas"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        Framebuscar.Visible = True
        DBGridBuscar.Columns(1).Width = "4000"
        TxtBuscar.SetFocus
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                BLinea = True
                BEquipo = False
                BDepartamento = False
                DataBuscar.RecordSource = "Select Linea, Descrip From Lineas"
                DataBuscar.Refresh
                DBGridBuscar.Refresh
                Framebuscar.Visible = True
                DBGridBuscar.Columns(1).Width = "4000"
                TxtBuscar.SetFocus
        End If
End Sub

Private Sub TxtTexto_Change()
        'EQUIPOS
        If OptOpcion.Item(0).Value = True Then
            Set RBusca = Db.OpenRecordset("Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtTexto.Text & "'")
                If RBusca.RecordCount > 0 Then
                    LblDes.Caption = RBusca!Descripcion
                Else
                    LblDes.Caption = ""
                End If
        'DEPARTAMENTOS
        ElseIf OptOpcion.Item(1).Value = True Then
            Set RBusca = Db.OpenRecordset("Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtTexto.Text & "'")
                If RBusca.RecordCount > 0 Then
                    LblDes.Caption = RBusca!Descripcion
                Else
                    LblDes.Caption = ""
                End If
        'EMPLEADOS
        ElseIf OptOpcion.Item(2).Value = True Then
            Set RBusca = Db.OpenRecordset("Select Descripcion From Empleados Where Codigo = '" & TxtTexto.Text & "'")
                If RBusca.RecordCount > 0 Then
                    LblDes.Caption = RBusca!Descripcion
                Else
                    LblDes.Caption = ""
                End If
        
        End If
End Sub

Private Sub TxtTexto_DblClick()
        If OptOpcion.Item(0).Value = True Then
            BEquipo = True
            BDepartamento = False
            BLinea = False
            BEmpleado = False
            DataBuscar.RecordSource = "Select Codigo, Descripcion From EmpleadosGrupos"
        ElseIf OptOpcion.Item(1).Value = True Then
            BEquipo = False
            BDepartamento = True
            BLinea = False
            BEmpleado = False
            DataBuscar.RecordSource = "Select Codigo, Descripcion From EmpleadosDepartamentos"
        ElseIf OptOpcion.Item(2).Value = True Then
            BEquipo = False
            BDepartamento = False
            BLinea = False
            BEmpleado = True
            DataBuscar.RecordSource = "Select Codigo, Descripcion From Empleados"
        End If
        
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
End Sub

Private Sub TxtTexto_GotFocus()
        TxtTexto.SelStart = 0
        TxtTexto.SelLength = Len(TxtTexto.Text)
End Sub

Private Sub TxtTexto_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                If OptOpcion.Item(0).Value = True Then
                    BEquipo = True
                    BDepartamento = False
                    BLinea = False
                    BEmpleado = False
                    DataBuscar.RecordSource = "Select Codigo, Descripcion From EmpleadosGrupos"
                ElseIf OptOpcion.Item(1).Value = True Then
                    BEquipo = False
                    BDepartamento = True
                    BLinea = False
                    BEmpleado = False
                    DataBuscar.RecordSource = "Select Codigo, Descripcion From EmpleadosDepartamentos"
                ElseIf OptOpcion.Item(2).Value = True Then
                    BEquipo = False
                    BDepartamento = False
                    BLinea = False
                    BEmpleado = True
                    DataBuscar.RecordSource = "Select Codigo, Descripcion From Empleados"
                End If
                
                    DataBuscar.Refresh
                    DBGridBuscar.Refresh
                    DBGridBuscar.Columns(1).Width = "4000"
                    Framebuscar.Visible = True
                    TxtBuscar.SetFocus
        End If
                
        
End Sub

Private Sub TxtTur_GotFocus()
        TxtTur.SelStart = 0
        TxtTur.SelLength = Len(TxtTur.Text)
End Sub

Private Sub TxtTur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
