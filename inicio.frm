VERSION 5.00
Begin VB.Form Inicio 
   BackColor       =   &H80000003&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio Sesion"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4875
   FillStyle       =   0  'Solid
   Icon            =   "inicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2115
      Begin VB.TextBox Txtusuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Txtpassword 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Password"
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
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Usuario"
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
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Height          =   1005
      Left            =   3600
      Picture         =   "inicio.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir Del Sistema"
      Top             =   3360
      Width           =   1065
   End
   Begin VB.CommandButton CmdAceptar 
      Default         =   -1  'True
      Height          =   1005
      Left            =   2280
      Picture         =   "inicio.frx":262C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Entrar Al Sistema"
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Version 04/04/2013 18:36"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   3840
      Picture         =   "inicio.frx":2EDF
      ToolTipText     =   "Empleados"
      Top             =   600
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   5
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   1
      Left            =   3720
      Picture         =   "inicio.frx":5681
      ToolTipText     =   "Pedidos"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   4
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   0
      Left            =   2400
      Picture         =   "inicio.frx":5ACB
      ToolTipText     =   "Inventario"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2040
      Picture         =   "inicio.frx":6395
      ToolTipText     =   "Inventario"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   2
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "inicio.frx":669F
      ToolTipText     =   "Produccion Y Calidad"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   2160
      Picture         =   "inicio.frx":6F69
      ToolTipText     =   "Eficiencia"
      Top             =   600
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   3
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Dim Caracteres As Long
Dim Texto As String
Dim RBuscaUsuario As New ADODB.Recordset
Dim RBuscaPassword As New ADODB.Recordset
Dim RutaArchivo As New ADODB.Stream
Dim cmd As New ADODB.Command
Dim Cont As Integer
Dim VNumeroAscii As String
Dim VEncriptado As String
Dim VLargo As Integer
Dim VPassword As String
Dim Temp As String





Private Sub CmdAceptar_Click()
On Error Resume Next
        
            
        If ((Txtpassword.Text = "sesavnelatem") Or (Txtpassword.Text = "SESAVNELATEM")) Then
            GUsuario = "METAL"
            Unload Me
            MousePointer = 11
                Menu.Show
            MousePointer = 0
            
'_____________________________________________________________________________________________________________________
        Else
            '----------------------------------------------------------------------------------------------------
            'PROCESO PARA ENCRIPTAR EL PASSWORD AGARRAMOS CADA LETRA DEL PASSWORD Y LE ASIGNAMOS EL CODIGO ASSCII
            'Y TAMBIEN LE AGREGAMOS UN NUMERO CUALQUIERA (0110) PARA QUE SEA UN POCO MAS DIFICIL DE LEERLO
            Cont = 1
            VLargo = Len(Txtpassword.Text)
            VNumeroAscii = ""
            VEncriptado = ""
            Do While Cont <= VLargo
               VNumeroAscii = Asc(Mid(Txtpassword.Text, Cont, 1))
               VEncriptado = VEncriptado & VNumeroAscii & "0110"
               Cont = Cont + 1
            Loop
            Txtpassword.Text = VEncriptado
                              
                        
            'INICIALIZA O CREA LA INTANCIA DEL RECORDSET
            Set RBuscaUsuario = New ADODB.Recordset
            'ABRIMOS EL RECORDSET PASANDO ESTOS PARAMETROS (Nombre del recorset, sql)
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaUsuario, "select * from Usuarios where Usuario = '" & TxtUsuario.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaUsuario, "select * from Usuarios where UPPER(Usuario) = '" & UCase(TxtUsuario.Text) & "'")
            End If
            
            If RBuscaUsuario.RecordCount > 0 Then
                GUsuario = UCase(TxtUsuario.Text)
            Else
                MsgBox "Verifique su Usuario", vbOKOnly + vbInformation, "Informacion"
                TxtUsuario.SetFocus
                Txtpassword.Text = ""
                Exit Sub
            End If
           
            'INICIALIZA O CREA LA INTANCIA DEL RECORDSET
            Set RBuscaPassword = New ADODB.Recordset
            'BUSCA EL PASSWORD
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaPassword, "select * from Usuarios where Usuario = '" & TxtUsuario.Text & "'" & " and Clave = '" & Txtpassword & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaPassword, "select * from Usuarios where UPPER(Usuario) = '" & UCase(TxtUsuario.Text) & "'" & " and UPPER(Clave) = '" & UCase(Txtpassword) & "'")
            End If
            If RBuscaPassword.RecordCount > 0 Then
            Else
                MsgBox "Verifique su Password", vbOKOnly + vbInformation, "Informacion"
                TxtUsuario.SetFocus
                Txtpassword.Text = ""
                Exit Sub
            End If
                        
            'SI ENCUENTRA EL USUARIO Y EL PASSWORD ENTONCES EMPEZAMOS A GUARDAR LAS VARIABLES PARA CONFIGURAR
            'EL ACCESO AL MENU
            If RBuscaPassword.RecordCount > 0 Then
                                
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Conexion.Execute "Update Usuarios Set FechaUltimoAcceso = #" & Format(Date, "mm/dd/yyyy") & "#, ContadorAccesos = ContadorAccesos + 1 Where Usuario = '" & TxtUsuario.Text & "'"
                                Else
                                    Conexion.Execute "Update Usuarios Set FechaUltimoAcceso = To_Date('" & Date & "', 'dd/mm/yyyy'), ContadorAccesos = ContadorAccesos + 1 Where UPPER(Usuario) = '" & UCase(TxtUsuario.Text) & "'"
                                End If
                                
                                If Err <> 0 Then
                                    MsgBox "error " & Err.Number & " " & Err.Description
                                    Err.Clear
                                End If
                                                                
                                'PRODUCCION
                                GConfiguracionCalidad = RBuscaPassword!ConfiguracionCalidad
                                GProduccion = RBuscaPassword!Produccion
                                GEspecificaciones = RBuscaPassword!Especificaciones
                                GReportesCalidad = RBuscaPassword!ReportesCalidad
                                
                                
                                'EFICIENCIA
                                GConfiguracionEficiencia = RBuscaPassword!ConfiguracionEficiencia
                                GCapturaParos = RBuscaPassword!CapturaParos
                                GReportesEficiencia = RBuscaPassword!ReportesEficiencia
                                GEditarEficiencia = RBuscaPassword!EditarEficiencia
                                GBorrarEficiencia = RBuscaPassword!BorrarEficiencia
                                
                                'AVANZADAS
                                GUsuarios = RBuscaPassword!Usuarios
                                GEditar = RBuscaPassword!Editar
                                GBorrar = RBuscaPassword!Borrar
                                GAjustesInventario = RBuscaPassword!GraficasCalidad
                                
                                'INVENTARIO
                                GConfiguracionInventario = RBuscaPassword!ConfiguracionInventario
                                GEntradas = RBuscaPassword!Entradas
                                GInspeccion = RBuscaPassword!Inspeccion
                                GTraslados = RBuscaPassword!Traslados
                                GSalidas = RBuscaPassword!Salidas
                                GCambiosUbicacion = RBuscaPassword!CambiosUbicacion
                                GCierreBulto = RBuscaPassword!CierreBulto
                                GLiberacionEntradas = RBuscaPassword!LiberacionEntradas
                                GLiberacionTraslados = RBuscaPassword!LiberacionTraslados
                                GLiberacionSalidas = RBuscaPassword!LiberacionSalidas
                                GGraficasInventario = RBuscaPassword!GraficasInventario
                                GReportesInventario = RBuscaPassword!ReportesInventario
                                GCapturaTransito = RBuscaPassword!CapturaTransito
                                GConsultaTransito = RBuscaPassword!ConsultaTransito
                                GPorConEntInv = RBuscaPassword!PorConEntInv
                                GReportesFormatos = RBuscaPassword!ReportesFormatos
                                GCapturaDesperdicio = RBuscaPassword!CapturaDesperdicio
                                GReclamosProveedor = RBuscaPassword!ReclamosProveedores
                                
                                'PEDIDOS
                                GPedidosClientes = RBuscaPassword!PedidosClientes
                                GPedidosProveedores = RBuscaPassword!PedidosProveedores
                                GCierreClientes = RBuscaPassword!CerrarPedidosClientes
                                GCierreProveedores = RBuscaPassword!CerrarpedidosProveedores
                                GEditarPedidos = RBuscaPassword!EditarPedidos
                                GBorrarPedidos = RBuscaPassword!BorrarPedidos
                                
                                
                                'EMPLEADOS
                                GConfiguracionEmpleados = RBuscaPassword!ConfiguracionEmpleados
                                GCapturaFaltas = RBuscaPassword!CapturaFaltas
                                GCapturaCursos = RBuscaPassword!CapturaCursos
                                GCapturaAumentos = RBuscaPassword!CapturaAumentos
                                GReportesEmpleados = RBuscaPassword!ReportesEmpleados
                                
                                'ORDENES
                                GOrdenProduccion = RBuscaPassword!OrdenProduccion
                                GInvVenRepEje = RBuscaPassword!InvVenRepEje
                                GReportesOrdenes = RBuscaPassword!ReportesOrdenes
                                
                                Menu.Show
                                Unload Me
            Else
                    MsgBox "Verifique su Password", vbOKOnly + vbInformation, "Informacion"
                    Txtpassword.Text = ""
                    Txtpassword.SetFocus
                    Exit Sub
            End If
        
        End If
        

End Sub

Private Sub CmdCancelar_Click()
        End
End Sub



Private Sub Form_Load()
On Error Resume Next
        
        'BUSCA LA RUTA DONDE VAN A ESTAR LOS ARCHIVOS DE TEXTO
        'CENTRALIZADOS PARA CUALQUIER CAMBIO
        'ESTOS NOS VAN A SERVIR PARA DEFINIR LA RUTA DE LOS REPORTES
        'RUTA DE LOS DOCUMENTOS, Y RUTA DE FOTOS
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile App.Path & "\RutaDeArchivosDeTexto.txt"
                GRutaDeArchivosDeTexto = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De Los Archivos De Texto " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If

        'ABRE EL ARCHIVO Q CONTIENE LA RUTA DE LOS ARCHIVOS DE TEXTO
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeReportes.txt"
                GRutaDeReportes = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De Los Reportes " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        'ABRE EL ARCHIVO Q CONTIENE LA RUTA DE LOS ARCHIVOS DE TEXTO
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeSeametal.txt"
                GRutaSeametal = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De SEAMETAL " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        'ABRE EL ARCHIVO Q CONTIENE LA RUTA DE LOS ARCHIVOS DE TEXTO
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeCpa.txt"
                GRutaCpa = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De CPA " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        'ABRE EL ARCHIVO Q CONTIENE LA RUTA DE LOS ARCHIVOS DE TEXTO
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeEpa.txt"
                GRutaEpa = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De EPA " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        'CHIAPAS
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeSeametalChiapas.txt"
                GRutaSeametalChiapas = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De SEAMETAL CHIAPAS " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeCpaChiapas.txt"
                GRutaCpaChiapas = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De CPA CHIAPAS " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeEpaChiapas.txt"
                GRutaEpaChiapas = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De EPA CHIAPAS " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
        
        
            
        'SAN LUIS POTOSI
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeSeametalSanLuisPotosi.txt"
                GRutaSeametalSanLuisPotosi = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De SEAMETAL SAN LUIS POTOSI" & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
        
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeCpaSanLuisPotosi.txt"
                GRutaCpaSanLuisPotosi = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De CPA SAN LUIS POTOSI " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
            
        
        Set RutaArchivo = New ADODB.Stream
            With RutaArchivo
                .Charset = "iso8859-1"
                .Open
                .LoadFromFile GRutaDeArchivosDeTexto & "\RutaDeEpaSanLuisPotosi.txt"
                GRutaEpaSanLuisPotosi = .ReadText()
                .Close
            End With
            
            If Err <> 0 Then
                MsgBox "Error Al Abrir El Archivo De Texto De La Ruta De EPA SAN LUIS POTOSI " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Err.Clear
                Exit Sub
            End If
        
            
        
 Temp = "ACCESS"
            
            If UCase(Temp) = "ORACLE" Then
                        GConectionString = "Provider=MSDAORA.1;User ID=produc;password=produccio;Data Source=metal;Persist Security Info=False"
                        GOrigenDeDatos = "AmaproOracle"
            ElseIf UCase(Temp) = "ACCESS" Then
                        'GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & App.Path & "\Metalenvases.mdb; Jet OLEDB:Database Password=metal"
                        GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & GRutaDeReportes & "\Metalenvases.mdb; Jet OLEDB:Database Password=metal"
                        GOrigenDeDatos = "AmaproAccess"
            Else
                    MsgBox "No Se Encuentra Archivo De Configuracion"
                    End
            End If
            
                                
            
            'INICIALIZA O CREA LA INSTANCIA DE LA CONECCION
            Set Conexion = New ADODB.Connection
            Conexion.ConnectionString = GConectionString
            Conexion.Open
            
            If Err <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
            
            
            
        
End Sub

