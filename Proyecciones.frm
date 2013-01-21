VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Proyecciones 
   Caption         =   "Proyecciones"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataProyecciones 
      Caption         =   "Proyecciones"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FichaTecnica"
      Top             =   7800
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGridProyecciones 
      Bindings        =   "Proyecciones.frx":0000
      Height          =   7575
      Left            =   120
      OleObjectBlob   =   "Proyecciones.frx":001F
      TabIndex        =   0
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "Proyecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
