VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menu Principal"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Repuestos"
      Height          =   1140
      Index           =   31
      Left            =   9720
      Picture         =   "Menu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ver Existencias Y Ubicaciones De Repuestos"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Rutinas TAPAS"
      Height          =   1020
      Index           =   18
      Left            =   240
      Picture         =   "Menu.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Consulta Actividades Diarias"
      Height          =   1140
      Index           =   15
      Left            =   8520
      Picture         =   "Menu.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Actividades Diarias"
      Height          =   1020
      Index           =   13
      Left            =   2640
      Picture         =   "Menu.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Histograma"
      Height          =   1020
      Index           =   10
      Left            =   1440
      Picture         =   "Menu.frx":211E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Transito"
      Height          =   1020
      Index           =   17
      Left            =   8520
      Picture         =   "Menu.frx":2560
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Producto En Transito"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Evaluacion Clientes"
      Height          =   1020
      Index           =   16
      Left            =   6120
      Picture         =   "Menu.frx":2E2A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Agrupar Lineas"
      Height          =   1140
      Index           =   29
      Left            =   4920
      Picture         =   "Menu.frx":36F4
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Evaluacion Proveedores"
      Height          =   1020
      Index           =   28
      Left            =   4920
      Picture         =   "Menu.frx":3FBE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Consulta Produccion Planta"
      Height          =   1140
      Index           =   26
      Left            =   7320
      Picture         =   "Menu.frx":4888
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Planificacion"
      Height          =   1020
      Index           =   24
      Left            =   8520
      Picture         =   "Menu.frx":5152
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Planificacion"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Tarimas Liberadas y No Cerradas"
      Height          =   1020
      Index           =   23
      Left            =   4920
      Picture         =   "Menu.frx":7FCC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Eficiencia"
      Height          =   1020
      Index           =   22
      Left            =   6240
      Picture         =   "Menu.frx":9CC6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Reportes"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Pedidos"
      Height          =   1020
      Index           =   20
      Left            =   8640
      Picture         =   "Menu.frx":A108
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Reportes"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Inventario"
      Height          =   1020
      Index           =   19
      Left            =   7440
      Picture         =   "Menu.frx":A41A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Reportes"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Batch a Produccion"
      Height          =   1020
      Index           =   14
      Left            =   6120
      Picture         =   "Menu.frx":ACE4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Inventario "
      Height          =   1020
      Index           =   3
      Left            =   9720
      Picture         =   "Menu.frx":10F6E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Consulta De Inventario"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Bulto/Tarima"
      Height          =   1020
      Index           =   12
      Left            =   9720
      Picture         =   "Menu.frx":11838
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Consulta De Bulto/Tarima"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Consulta Produccion Reporte"
      Height          =   1140
      Index           =   11
      Left            =   6120
      Picture         =   "Menu.frx":11B42
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Produccion"
      Height          =   1020
      Index           =   9
      Left            =   5040
      Picture         =   "Menu.frx":1240C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Reportes"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Produccion Liberada"
      Height          =   1020
      Index           =   2
      Left            =   240
      Picture         =   "Menu.frx":12716
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Rutinas CPA"
      Height          =   1020
      Index           =   6
      Left            =   1440
      Picture         =   "Menu.frx":12A20
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Rutinas SEAMETAL"
      Height          =   1020
      Index           =   5
      Left            =   240
      Picture         =   "Menu.frx":12D2A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Genera &Batch y Certificado"
      Height          =   1020
      Index           =   8
      Left            =   1440
      Picture         =   "Menu.frx":13034
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Genera Rutinas"
      Height          =   1020
      Index           =   7
      Left            =   240
      Picture         =   "Menu.frx":138FE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Captura &Rutinas"
      Height          =   1020
      Index           =   4
      Left            =   240
      Picture         =   "Menu.frx":13D40
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Producción "
      Height          =   1020
      Index           =   1
      Left            =   240
      Picture         =   "Menu.frx":14182
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Lineas"
      Height          =   1020
      Index           =   0
      Left            =   240
      Picture         =   "Menu.frx":1448C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   4395
      Index           =   1
      Left            =   4800
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "www.fepsa.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      ToolTipText     =   "Click para entrar a Intranet"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   4920
      TabIndex        =   25
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Shape Shape1 
      Height          =   1275
      Index           =   0
      Left            =   4920
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Reportes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   0
      Left            =   5040
      TabIndex        =   22
      Top             =   240
      Width           =   1275
   End
   Begin VB.Image ImgFondo 
      Height          =   8895
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Menu Produccion 
      Caption         =   "Calidad Y Produccion"
      Begin VB.Menu MnuConfiguracionCalidad 
         Caption         =   "Configuracion"
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Tipos De Ficha Tecnica"
            Index           =   0
         End
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Ficha Tecnica "
            Index           =   1
         End
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Asignacion De Materias Primas"
            Index           =   2
         End
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Turnos"
            Index           =   26
         End
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Lineas Personal Turno"
            Index           =   28
         End
         Begin VB.Menu MnuMantenimientos 
            Caption         =   "Tipos De Ficha Tecnica VENTAS"
            Index           =   29
         End
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Lineas"
         Index           =   1
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Captura De Produccion"
         Index           =   2
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Captura De Materias Primas En Produccion"
         Index           =   3
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Captura De Produccion Liberada"
         Index           =   4
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Captura de Rutinas"
         Index           =   6
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Genera Rutinas"
         Index           =   9
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Genera Batch"
         Index           =   10
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Genera Captura Rutinas Automatica SEAMETAL 9000"
         Index           =   11
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "Genera Captura Rutinas Automatica CPA"
         Index           =   12
      End
      Begin VB.Menu MnuProduccion 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuEspecificaciones 
         Caption         =   "Catalogos Especificaciones"
      End
      Begin VB.Menu MnuRutinas 
         Caption         =   "Rutinas"
      End
      Begin VB.Menu MnuDefectos 
         Caption         =   "Defectos"
      End
      Begin VB.Menu MnuAtributos 
         Caption         =   "Atributos"
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReportes 
         Caption         =   "Reportes"
      End
      Begin VB.Menu MnuGraficas 
         Caption         =   "Graficas Y Consultas"
         Begin VB.Menu MnuGeneraHistograma 
            Caption         =   "Grafica De Histograma"
         End
         Begin VB.Menu MnuGerencia 
            Caption         =   "Consulta De Produccion"
         End
      End
      Begin VB.Menu mnucamcab 
         Caption         =   "Actualizar Paros CF y MP"
      End
   End
   Begin VB.Menu MnuOrdenes 
      Caption         =   "Ordenes"
      Begin VB.Menu MnuOrden 
         Caption         =   "Ordenes De Produccion"
         Index           =   0
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Pasadas De Barniz"
         Index           =   1
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Genera Inventario Automatico"
         Index           =   3
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Genera Ventas Automatico"
         Index           =   4
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Inventario"
         Index           =   6
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Ventas"
         Index           =   7
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Ventas Metas Mensuales"
         Index           =   8
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuOrden 
         Caption         =   "Reportes"
         Index           =   10
      End
      Begin VB.Menu MnuOrden 
         Caption         =   ""
         Index           =   11
      End
   End
   Begin VB.Menu MenuProduccion 
      Caption         =   "Eficiencia"
      Begin VB.Menu MnuConfiguracionEficiencia 
         Caption         =   "Configuracion"
         Begin VB.Menu MnuEfiCon 
            Caption         =   "Paros Grupos"
            Index           =   0
         End
         Begin VB.Menu MnuEfiCon 
            Caption         =   "Paros"
            Index           =   1
         End
         Begin VB.Menu MnuEfiBla 
            Caption         =   ""
         End
      End
      Begin VB.Menu MnuCapturaParos 
         Caption         =   "Captura de Paros"
      End
      Begin VB.Menu MnuReportesEficiencia 
         Caption         =   "Reportes (eficiencia)"
      End
      Begin VB.Menu UltimaLineaProduccion 
         Caption         =   ""
      End
   End
   Begin VB.Menu MnuInventarioMateriaPrima 
      Caption         =   "Inventarios"
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Configuracion"
         Index           =   0
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Bodegas Grupos"
            Index           =   0
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Bodegas"
            Index           =   1
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Clientes"
            Index           =   4
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Proveedores Grupos"
            Index           =   5
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Proveedores"
            Index           =   6
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Transportistas"
            Index           =   7
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Procesos De Desperdicio"
            Index           =   8
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Tipos De Entradas"
            Index           =   9
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Tipos De Documentos"
            Index           =   10
         End
         Begin VB.Menu MnuMateriaPrimaConfiguracion 
            Caption         =   "Unidades De Medida"
            Index           =   11
         End
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Entradas"
         Index           =   2
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Inspeccion"
         Index           =   3
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Traslados"
         Index           =   4
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Salidas"
         Index           =   5
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Liberacion Entradas"
         Index           =   7
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Liberacion Traslados"
         Index           =   8
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Liberacion Salidas"
         Index           =   9
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Cerrar Bulto/Tarima"
         Index           =   11
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Captura De Desperdicio"
         Index           =   13
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Reportes"
         Index           =   15
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Reportes Formatos Inspeccion"
         Index           =   16
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "% Conforme Pedidos Proveedores"
         Index           =   18
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Cobros De Reclamos A Proveedor"
         Index           =   19
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Captura Producto En Transito"
         Index           =   20
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Consulta Producto Transito"
         Index           =   21
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Cambios Ubicacion"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   "Consultas"
         Index           =   23
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMateriaPrima 
         Caption         =   ""
         Index           =   24
      End
   End
   Begin VB.Menu MenuPedidos 
      Caption         =   "Pedidos"
      Begin VB.Menu MnuPedidos 
         Caption         =   "Pedidos A Proveedores"
         Index           =   0
      End
      Begin VB.Menu MnuPedidos 
         Caption         =   "Pedidos De Clientes"
         Index           =   1
      End
      Begin VB.Menu MnuPedidos 
         Caption         =   "Cierre Pedidos A Proveedores"
         Index           =   2
      End
      Begin VB.Menu MnuPedidos 
         Caption         =   "Cierre Pedidos De Clientes"
         Index           =   3
      End
      Begin VB.Menu MnuPedidos 
         Caption         =   "Reportes"
         Index           =   5
      End
      Begin VB.Menu MnuPedidos 
         Caption         =   ""
         Index           =   6
      End
   End
   Begin VB.Menu MnuEmp 
      Caption         =   "Empleados"
      Begin VB.Menu MnuEmpConfiguracion 
         Caption         =   "Configuracion"
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Departamentos"
            Index           =   0
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Equipos"
            Index           =   1
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Puestos"
            Index           =   2
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Escolaridades"
            Index           =   3
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Faltas"
            Index           =   5
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Cursos"
            Index           =   6
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Habilidades"
            Index           =   8
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Habilidades De Empleados"
            Index           =   9
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Habilidades De Puestos"
            Index           =   10
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Empleados"
            Index           =   14
         End
         Begin VB.Menu MnuEmpCon 
            Caption         =   "Hijos "
            Index           =   15
         End
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Genera Horas Automatico"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Captura De Horas"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Captura Faltas"
         Index           =   4
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Captura Cursos"
         Index           =   5
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Captura Aumentos"
         Index           =   6
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Captura Vacaciones"
         Index           =   7
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   "Reportes"
         Index           =   9
      End
      Begin VB.Menu MnuEmpleados 
         Caption         =   ""
         Index           =   10
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Avanzadas"
      Begin VB.Menu MenuAccesos 
         Caption         =   "Usuarios"
         Index           =   0
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "Modifica Bultos/Tarimas"
         Index           =   3
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "Ajustes De Inventario"
         Index           =   6
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   "Borra Actividades Diarias"
         Index           =   8
      End
      Begin VB.Menu MenuAccesos 
         Caption         =   ""
         Index           =   9
      End
   End
   Begin VB.Menu MnuCambiarPassword 
      Caption         =   "Cambiar Password"
   End
   Begin VB.Menu MnuSalida 
      Caption         =   "Salida"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim r As New ADODB.Recordset
'Dim r2 As New ADODB.Recordset
Dim VTexto As String


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
        MousePointer = 11
        'LINEAS
        If Index = 0 Then
                Lineas.Show
        'CAPTURA DE PRODUCCION
        ElseIf Index = 1 Then
                CapturaProduccion.Show
        'CAPTURA DE PRODUCCION LIBERADA
        ElseIf Index = 2 Then
                CapturaProduccionLiberada.Show
        'CAPTURA DE RUTINAS
        ElseIf Index = 4 Then
                CapturaRutinas.Show
        'CAPTURA DE RUTINAS SEA INTERNA
        ElseIf Index = 5 Then
                ElegirPlanta.Show 1
                'GeneraCapturaRutinas.Show
        'CAPTURA DE RUTINAS CPA
        ElseIf Index = 6 Then
                ElegirPlanta2.Show 1
                GeneraCapturaRutinasCpa.Show
        'GENERA LA CAPTURA DE RUTINAS
        ElseIf Index = 7 Then
                GeneraRutinas.Show
        'GENERA BATCH
        ElseIf Index = 8 Then
                GeneraBatch.Show
        'REPORTES
        ElseIf Index = 9 Then
                Reportes.Show
        ElseIf Index = 10 Then
                GeneraHistograma.Show
        'CONSULTA ESPECIAL DE INVENTARIO DE PRODUCTO TERMINADO
        ElseIf Index = 3 Then
                ConsultaInventarioProductoTerminado.Show
        'CONSULTA ESPECIAL DE PRODUCCION
        ElseIf Index = 11 Then
                'Gerencia.Show
                ConsultaDeProduccionReporte.Show
        'CONSULTA ESPECIAL DE BULTO MATERIA PRIMA
        'CONSULTA ESPECIAL DE TARIMA PRODUCTO TERMINADO
        ElseIf Index = 12 Then
                ConsultaTarima.Show
        ElseIf Index = 13 Then
                'ActividadesDiarias.Show
        'BATCH NO INGRESADOS A INVENTARIO
        ElseIf Index = 14 Then
                BatchNoIngresadosAInventario.Show
        ElseIf Index = 15 Then
                'ActividadesDiariasConsulta.Show
        'CONSULTA DESPACHOS PROVEEDOR
        ElseIf Index = 17 Then
               ControlDeDespachosConsulta.Show
        'RUTINAS TAPAS
        ElseIf Index = 18 Then
               ElegirPlanta3.Show
        'REPORTES PRODUCTO TERMINADO
        ElseIf Index = 19 Then
                ReportesProductoTerminado.Show
        'REPORTES PEDIDOS
        ElseIf Index = 20 Then
                ReportesPedidos.Show
        'REPORTES EFICIENCIA
        ElseIf Index = 22 Then
                ReportesDeEficiencia.Show
        'TARIMAS LIBERADAS Y NO CERRADAS
        ElseIf Index = 23 Then
                TarimasLiberadasNoCerradas.Show
        'PLANIFICACION
        ElseIf Index = 24 Then
                Planificacion.Show
        'GRAFICA DE DEFECTOS
        ElseIf Index = 25 Then
                
        'CONSULTA DE PRODUCCION DE PLANTA
        ElseIf Index = 26 Then
                ConsultaDeProduccionCalidad.Show
        'HISTOGRAMA
        ElseIf Index = 27 Then
            GeneraHistograma.Show
        'EVALUACION DE PROVEEDORES
        ElseIf Index = 28 Then
            EvaluacionProveedores.Show
        'EVALUACION DE CLIENTES
        ElseIf Index = 16 Then
            EvaluacionClientes.Show
        'AGRUPAR LINEAS
        ElseIf Index = 29 Then
            AgruparLineas.Show
        ElseIf Index = 31 Then
            RepuestosConsulta.Show
        
        End If
        
        
        
        If Err <> 0 And Err <> 380 Then
            'MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
            Err.Clear
        End If
        MousePointer = 0
End Sub









Private Sub Form_Load()
On Error Resume Next
    'CARGA LA PORTADA
    ImgFondo.Picture = LoadPicture(GRutaDeReportes & "\fondo.jpg")
    
    
    If Err <> 0 Then
    End If
        
    Menu.Caption = "Menu Principal Evadeva Y Amapro " & Space(75) & "Usuario " & GUsuario & " " & Time
    
If GUsuario = "METAL" Then
Else

        'PRODUCCION ____________________________________________________________________________
                    If GConfiguracionCalidad = False Then
                        MnuConfiguracionCalidad.Visible = False
                    End If
                    'MENU DE PRODUCCION
                    If GProduccion = False Then
                            MnuProduccion.Item(1).Visible = False
                            MnuProduccion.Item(2).Visible = False
                            MnuProduccion.Item(3).Visible = False
                            MnuProduccion.Item(4).Visible = False
                            MnuProduccion.Item(5).Visible = False
                            MnuProduccion.Item(6).Visible = False
                            MnuProduccion.Item(7).Visible = False
                            MnuProduccion.Item(8).Visible = False
                            MnuProduccion.Item(9).Visible = False
                            MnuProduccion.Item(10).Visible = False
                            MnuProduccion.Item(11).Visible = False
                            MnuProduccion.Item(12).Visible = False
                            MnuProduccion.Item(13).Visible = False
                            MnuProduccion.Item(14).Visible = False
                            MnuProduccion.Item(15).Visible = False
                            
                            CmdBotones.Item(0).Visible = False
                            CmdBotones.Item(1).Visible = False
                            CmdBotones.Item(2).Visible = False
                            CmdBotones.Item(4).Visible = False
                            CmdBotones.Item(5).Visible = False
                            CmdBotones.Item(6).Visible = False
                            CmdBotones.Item(7).Visible = False
                            CmdBotones.Item(8).Visible = False
                            
                            
                    End If
                    'MENU DE ESPECIFICACIONES
                    If GEspecificaciones = False Then
                            MnuEspecificaciones.Visible = False
                            MnuRutinas.Visible = False
                            MnuDefectos.Visible = False
                            MnuAtributos.Visible = False
                            
                    End If
                    'MENU DE REPORTES
                    If GReportesCalidad = False Then
                            MnuReportes.Visible = False
                            CmdBotones.Item(9).Visible = False
                    End If
                    
                    
        'EFICIENCIA ____________________________________________________________________________
                    
                    If GConfiguracionEficiencia = False Then
                          MnuConfiguracionEficiencia.Visible = False
                    End If
                    If GCapturaParos = False Then
                            MnuCapturaParos.Visible = False
                    End If
                    If GReportesEficiencia = False Then
                            MnuReportesEficiencia.Visible = False
                    End If
                    
        'ORDENES DE PRODUCCION Y PASADAS
                    If GOrdenProduccion = False Then
                        MnuOrden.Item(0).Visible = False
                        MnuOrden.Item(1).Visible = False
                    End If
                    
                    'INVENTARIO VENTAS Y REPORTE EJECUTIVO
                    If GInvVenRepEje = False Then
                        MnuOrden.Item(3).Visible = False
                        MnuOrden.Item(4).Visible = False
                        MnuOrden.Item(6).Visible = False
                        MnuOrden.Item(7).Visible = False
                    End If
                    'REPORTES
                    If GReportesOrdenes = False Then
                        MnuOrden.Item(9).Visible = False
                    End If
                    
                    
                    
        
        'INVENTARIO ________________________________________________________________________
                                        
                    If GConfiguracionInventario = False Then
                        MnuMateriaPrima.Item(0).Visible = False
                    End If
                    If GEntradas = False Then
                        MnuMateriaPrima.Item(2).Visible = False
                    End If
                    If GInspeccion = False Then
                        MnuMateriaPrima.Item(3).Visible = False
                    End If
                    If GTraslados = False Then
                        MnuMateriaPrima.Item(4).Visible = False
                    End If
                    If GSalidas = False Then
                        MnuMateriaPrima.Item(5).Visible = False
                    End If
                    If GCambiosUbicacion = False Then
                        MnuMateriaPrima.Item(22).Visible = False
                    End If
                    If GCierreBulto = False Then
                        MnuMateriaPrima.Item(11).Visible = False
                    End If
                    If GLiberacionEntradas = False Then
                        MnuMateriaPrima.Item(7).Visible = False
                    End If
                    If GLiberacionTraslados = False Then
                        MnuMateriaPrima.Item(8).Visible = False
                    End If
                    If GLiberacionSalidas = False Then
                        MnuMateriaPrima.Item(9).Visible = False
                    End If
                    If GGraficasInventario = False Then
                        MnuMateriaPrima.Item(23).Visible = False
                    End If
                    If GReportesInventario = False Then
                        MnuMateriaPrima.Item(15).Visible = False
                    End If
                    If GCapturaTransito = False Then
                        MnuMateriaPrima.Item(20).Visible = False
                    End If
                    If GConsultaTransito = False Then
                        MnuMateriaPrima.Item(21).Visible = False
                        CmdBotones.Item(17).Visible = False
                    End If
                    If GPorConEntInv = False Then
                        MnuMateriaPrima.Item(18).Visible = False
                    End If
                    If GReportesFormatos = False Then
                        MnuMateriaPrima.Item(16).Visible = False
                    End If
                    If GCapturaDesperdicio = False Then
                        MnuMateriaPrima.Item(13).Visible = False
                    End If
                    If GReclamosProveedor = False Then
                        MnuMateriaPrima.Item(19).Visible = False
                    End If
                    
                    
            
        'PEDIDOS ______________________________________________________________________
                    If GPedidosClientes = False Then
                        MnuPedidos.Item(1).Visible = False
                    End If
                    If GPedidosProveedores = False Then
                        MnuPedidos.Item(0).Visible = False
                    End If
                    If GCierreClientes = False Then
                        MnuPedidos.Item(3).Visible = False
                    End If
                    If GCierreProveedores = False Then
                        MnuPedidos.Item(2).Visible = False
                    End If
        
        'EMPLEADOS ______________________________________________________________________
                    If GConfiguracionEmpleados = False Then
                        MnuEmpConfiguracion.Visible = False
                    End If
                    'If GEmpleadosGeneraHoras = False Then
                        'MnuEmpleados.Item(1).Visible = False
                    'End If
                    'If GEmpleadosCapturaHoras = False Then
                        'MnuEmpleados.Item(2).Visible = False
                    'End If
                    If GCapturaFaltas = False Then
                        MnuEmpleados.Item(4).Visible = False
                    End If
                    If GCapturaCursos = False Then
                        MnuEmpleados.Item(5).Visible = False
                    End If
                    If GCapturaAumentos = False Then
                        MnuEmpleados.Item(6).Visible = False
                    End If
                    If GReportesEmpleados = False Then
                        MnuEmpleados.Item(8).Visible = False
                    End If
                    
        'USUARIOS ______________________________________________________________________
        
                    If GUsuarios = False Then
                            
                            MenuAccesos.Item(0).Visible = False
                            MenuAccesos.Item(3).Visible = False
                    End If
                    
                    'AJUSTES
                    If GAjustesInventario = False Then
                            MenuAccesos.Item(6).Visible = False
                    End If
End If
            
End Sub


Private Sub Form_Resize()
On Error Resume Next
    ImgFondo.Top = 0
    ImgFondo.Left = 0
    ImgFondo.Height = ScaleHeight
    ImgFondo.Width = ScaleWidth
    If Err <> 0 Then
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
        Desconectar
        
'Dim gotoval As Integer
'Dim gointo As Integer

        'cerrar el formulaio
'gotoval = Me.Height / 2
'For gointo = 1 To gotoval
'    DoEvents
'        Me.Height = Me.Height - 10
'        If Me.Height <= 11 Then GoTo horiz
'    Next gointo
'horiz:
'    Me.Height = 30
'    gotoval = Me.Width / 2
'    For gointo = 1 To gotoval
'        DoEvents
'            Me.Width = Me.Width - 10
'            If Me.Width <= 11 Then End
''        Next gointo

    If Err <> 0 Then
    End If

End Sub

Private Sub GraficaDEDefectos_Click()
        'GraficaDefectos.Show
End Sub

Private Sub Label1_Click()
        Internet.Show
End Sub

Private Sub MenuAccesos_Click(Index As Integer)
    If Index = 0 Then
        Usuarios.Show
    ElseIf Index = 3 Then
        InventarioModificaBultos.Show
    ElseIf Index = 6 Then
        AjustesProductoTerminado.Show
    ElseIf Index = 8 Then
        'ActividadesDiariasBorrar.Show
    End If
End Sub


Private Sub MnuAtributos_Click()
        Atributos.Show
End Sub

Private Sub MnuCambiarPassword_Click()
        CambiarClave.Show
End Sub

Private Sub mnucamcab_Click()
' ActualizaParos.Show
End Sub

Private Sub MnuCapturaParos_Click()
        MousePointer = 11
            CapturaParos.Show
        MousePointer = 0
End Sub

Private Sub MnuDefectos_Click()
    Defectos.Show
End Sub

Private Sub MnuEfiCon_Click(Index As Integer)
        If Index = 0 Then
                ParosGrupos.Show
        ElseIf Index = 1 Then
                Paros.Show
        End If
        
End Sub


Private Sub MnuEmpCon_Click(Index As Integer)
        If Index = 0 Then
            EmpleadosDepartamentos.Show
        ElseIf Index = 1 Then
            EmpleadosGrupos.Show
        ElseIf Index = 2 Then
            EmpleadosPuestos.Show
        ElseIf Index = 3 Then
            EmpleadosEscolaridad.Show
        ElseIf Index = 4 Then
            'LINEA
        ElseIf Index = 5 Then
            EmpleadosFaltas.Show
        ElseIf Index = 6 Then
            EmpleadosCursos.Show
        ElseIf Index = 7 Then
            'LINEA
        ElseIf Index = 8 Then
            EmpleadosHabilidades.Show
        ElseIf Index = 9 Then
            EmpleadosHabilidadesEmpleado.Show
        ElseIf Index = 10 Then
            EmpleadosHabilidadesPuesto.Show
        ElseIf Index = 11 Then
            'LINEA
        ElseIf Index = 13 Then
            'LINEA
        ElseIf Index = 14 Then
            Empleados.Show
        ElseIf Index = 15 Then
            EmpleadosHijos.Show
        End If
End Sub


Private Sub MnuEmpleados_Click(Index As Integer)
        If Index = 0 Then
        
        ElseIf Index = 1 Then
            'EmpleadosHorasAutomatico.Show
        ElseIf Index = 2 Then
           ' EmpleadosCapturaHoras.Show
        ElseIf Index = 4 Then
            EmpleadosCapturaFaltas.Show
        ElseIf Index = 5 Then
            EmpleadosCapturaCursos.Show
        ElseIf Index = 6 Then
            EmpleadosCapturaAumentos.Show
        ElseIf Index = 7 Then
            EmpleadosCapturaVacaciones.Show
        ElseIf Index = 8 Then
            'LINEA
        ElseIf Index = 9 Then
            ReportesDeEmpleados.Show
        End If
End Sub

Private Sub MnuEspecificaciones_Click()
    MousePointer = 11
            CatalogosEspecificaciones.Show
    MousePointer = 0
End Sub


Private Sub MnuGeneraHistograma_Click()
    MousePointer = 11
        GeneraHistograma.Show
    MousePointer = 0
End Sub
Private Sub MnuGerencia_Click()
        ConsultaDeProduccionCalidad.Show
End Sub

Private Sub MnuOrden_Click(Index As Integer)
                If Index = 0 Then
                    OrdenProduccion.Show
                ElseIf Index = 1 Then
                    Pasadas.Show
                ElseIf Index = 2 Then
                    'LINEA
                ElseIf Index = 3 Then
                    GeneraInventario.Show
                ElseIf Index = 4 Then
                    GeneraVentas.Show
                ElseIf Index = 5 Then
                    'LINEA
                ElseIf Index = 6 Then
                    Inventario.Show
                ElseIf Index = 7 Then
                    Ventas.Show
                ElseIf Index = 8 Then
                    VentasMetas.Show
                ElseIf Index = 10 Then
                    ReportesDeOrdenes.Show
                End If
End Sub

Private Sub MnuPedidos_Click(Index As Integer)
                If Index = 0 Then
                    PedidosProveedores.Show
                ElseIf Index = 1 Then
                    PedidosClientes.Show
                ElseIf Index = 2 Then
                    CierrePedidosProveedores.Show
                ElseIf Index = 3 Then
                    CierrePedidosClientes.Show
                ElseIf Index = 5 Then
                    ReportesPedidos.Show
                End If
End Sub


Private Sub MnuMantenimientos_Click(Index As Integer)
On Error Resume Next

MousePointer = 11
                If Index = 0 Then
                        FichaTecnicaTipos.Show
                ElseIf Index = 1 Then
                        FichaTecnica.Show
                ElseIf Index = 2 Then
                        FichaTecnicaConMateriaPrima.Show
                ElseIf Index = 26 Then
                        Turnos.Show 1
                ElseIf Index = 28 Then
                        LineasPersonalTurno.Show
                ElseIf Index = 29 Then
                        FichaTecnicaTiposVentas.Show
                End If
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
                End If
MousePointer = 0
                
End Sub

Private Sub MnuMateriaPrima_Click(Index As Integer)
On Error Resume Next
        MousePointer = 11
                If Index = 0 Then
                    'MENU DE CONFIGURACION
                ElseIf Index = 1 Then
                    'LINEA
                ElseIf Index = 2 Then
                    InventarioEntradas.Show
                ElseIf Index = 3 Then
                    InventarioInspeccion.Show
                ElseIf Index = 4 Then
                    InventarioTraslados.Show
                ElseIf Index = 5 Then
                    InventarioSalidas.Show
                ElseIf Index = 7 Then
                    InventarioLiberacionEntradas.Show
                ElseIf Index = 8 Then
                    InventarioLiberacionTraslados.Show
                ElseIf Index = 9 Then
                    InventarioLiberacionSalidas.Show
                ElseIf Index = 11 Then
                    CerrarBulto.Show 1
                ElseIf Index = 13 Then
                    CapturaDesperdicio.Show
                ElseIf Index = 15 Then
                    ReportesProductoTerminado.Show
                ElseIf Index = 16 Then
                    ReportesFormatos.Show
                ElseIf Index = 18 Then
                    PorcentajeNoConforme.Show
                ElseIf Index = 19 Then
                    CobrosProveedor.Show
                ElseIf Index = 20 Then
                    ControlDeDespachos.Show
                ElseIf Index = 21 Then
                    ControlDeDespachosConsulta.Show
                End If
                
        If Err <> 0 Then
            'MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
            'Exit Sub
        End If
        
        MousePointer = 0
End Sub

Private Sub MnuMateriaPrimaConfiguracion_Click(Index As Integer)
On Error Resume Next
                If Index = 0 Then
                        BodegasInventarioGrupos.Show
                ElseIf Index = 1 Then
                        BodegasInventario.Show
                ElseIf Index = 4 Then
                        Clientes.Show
                ElseIf Index = 5 Then
                        ProveedoresGrupos.Show
                ElseIf Index = 6 Then
                        Proveedores.Show
                ElseIf Index = 7 Then
                        Transportistas.Show
                ElseIf Index = 8 Then
                        ProcesosMateriaPrima.Show
                ElseIf Index = 9 Then
                        TiposEntradaMateriaPrima.Show
                ElseIf Index = 10 Then
                        Documentos.Show
                ElseIf Index = 11 Then
                        UnidadesMedida.Show
                End If
                
                If Err <> 0 Then
                    'MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
                    'Exit Sub
                End If
End Sub

Private Sub MnuProduccion_Click(Index As Integer)
On Error Resume Next

MousePointer = 11
                If Index = 1 Then
                        Lineas.Show
                ElseIf Index = 2 Then
                        CapturaProduccion.Show
                ElseIf Index = 3 Then
                        CapturaProduccionMateriaPrima.Show
                ElseIf Index = 4 Then
                        CapturaProduccionLiberada.Show
                ElseIf Index = 6 Then
                        CapturaRutinas.Show
                ElseIf Index = 7 Then
                        'LINEA
                ElseIf Index = 8 Then
                        'ESPACIO
                ElseIf Index = 9 Then
                        GeneraRutinas.Show
                ElseIf Index = 10 Then
                        GeneraBatch.Show
                ElseIf Index = 11 Then
                        GeneraCapturaRutinas.Show
                ElseIf Index = 12 Then
                        GeneraCapturaRutinasCpa.Show
                ElseIf Index = 13 Then
                        'LINEA
                ElseIf Index = 14 Then
                
                End If
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
                    Exit Sub
                End If
MousePointer = 0
End Sub


Private Sub MnuRepForIns_Click()
            MousePointer = 11
                ReportesFormatos.Show
            MousePointer = 0
End Sub


Private Sub MnuReportes_Click()
            MousePointer = 11
                Reportes.Show
            MousePointer = 0
End Sub

Private Sub MnuReportesEficiencia_Click()
            MousePointer = 11
                ReportesDeEficiencia.Show
            MousePointer = 0
End Sub

Private Sub MnuRutinas_Click()
    Rutinas.Show
End Sub

Private Sub MnuSalida_Click()
'On Error Resume Next

'Dim gotoval As Integer
'Dim gointo As Integer

        'cerrar el formulaio
'gotoval = Me.Height / 2
'For gointo = 1 To gotoval
'    DoEvents
'        Me.Height = Me.Height - 10
'        If Me.Height <= 11 Then GoTo horiz
'    Next gointo
'horiz:
'    Me.Height = 30
'    gotoval = Me.Width / 2
'    For gointo = 1 To gotoval
'        DoEvents
'            Me.Width = Me.Width - 10
'            If Me.Width <= 11 Then End
'        Next gointo'

'If Err <> 0 Then

'End If

    End 'sale
    
End Sub


Private Sub UltimaLinea_Click()
'    actualizallavecapturarutinas.Show
End Sub
