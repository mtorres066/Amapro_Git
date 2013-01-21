VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentasNaf 
   Caption         =   "Ventas De Naf"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGridVentas 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9340
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton Generar 
      Caption         =   "&Generar"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "VentasNaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RVentas As ADODB.Recordset

Private Sub Form_Load()
           'ACTIVA LA VARIABLE CONEXION QUE ES TIPO BASE DE DATOS
           Crear_instancias
           'Call Param_Conexion("msdasql", TxtUsuario.Text, TxtPassword.Text)
           Call Param_Conexion("Naf", Ealvarez, planta)
End Sub

Private Sub Generar_Click()
           Set RVentas = New ADODB.Recordset
           
           Call Abrir_Recordset(RVentas, "Select * From EncaFacturass")
           
           Set DataGridVentas = RVentas
                      
End Sub
