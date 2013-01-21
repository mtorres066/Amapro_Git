VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BatchNoIngresadosAInventario 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch No Entregados A Inventario"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "BatchNoIngresadosAInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGridBatch2 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
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
      Caption         =   "Produccion Liberada"
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
   Begin MSDataGridLib.DataGrid DataGridBatch 
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9128
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
      Caption         =   "Produccion"
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
   Begin MSComCtl2.DTPicker DTPMes 
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MM/yyyy"
      Format          =   61407235
      CurrentDate     =   37578
   End
   Begin VB.CommandButton CmdBotones 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   3240
      Picture         =   "BatchNoIngresadosAInventario.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salida"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CmdBotones 
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   2640
      Picture         =   "BatchNoIngresadosAInventario.frx":4814
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Consultar"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Mes Y Año"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   930
   End
End
Attribute VB_Name = "BatchNoIngresadosAInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RProduccion As New ADODB.Recordset
Dim RProduccionLiberada As New ADODB.Recordset



Private Sub CmdBotones_Click(Index As Integer)
    
    'CONSULTA
    If Index = 0 Then
        MousePointer = 11
           'PRODUCCION
           Set RProduccion = New ADODB.Recordset
           If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RProduccion, "SELECT DISTINCTROW P.Linea, P.Batch FROM Produccion as P LEFT JOIN DetalleEntradasinventario As DE ON P.Linea = DE.Linea And P.Batch = DE.Batch Where (((DE.Batch) Is Null)) And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value) & " Group By P.Linea, P.Batch")
           Else 'ORACLE
                Call Abrir_Recordset(RProduccion, "SELECT DISTINCT P.Linea, P.Batch FROM Produccion P LEFT JOIN DetalleEntradasInventario DE ON P.Linea = DE.Linea And P.Batch = DE.Batch Where (((DE.Batch) Is Null)) And To_Char(P.Fec_Prd,'mm') = '" & Format(DTPMes.Value, "mm") & "' And To_Char(P.Fec_Prd,'yyyy') = '" & Year(DTPMes.Value) & "' Group By P.Linea, P.Batch")
           End If
           Set DataGridBatch.DataSource = RProduccion
           
           
           'BATCH DE PRODUCCION LIBERADA
           Set RProduccionLiberada = New ADODB.Recordset
           If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RProduccionLiberada, "SELECT DISTINCTROW P.Linea, P.Batch FROM ProduccionLiberada as P LEFT JOIN DetalleEntradasInventario As DE ON P.Linea = DE.Linea And P.Batch = DE.Batch Where (((DE.Batch) Is Null)) And Month(P.Fec_Prd) = " & Month(DTPMes.Value) & " And Year(P.Fec_Prd) = " & Year(DTPMes.Value) & " Group By P.Linea, P.Batch")
           Else 'ORACLE
                Call Abrir_Recordset(RProduccionLiberada, "SELECT DISTINCT P.Linea, P.Batch FROM ProduccionLiberada P LEFT JOIN DetalleEntradasInventario DE ON P.Linea = DE.Linea And P.Batch = DE.Batch Where (((DE.Batch) Is Null)) And To_Char(P.Fec_Prd,'mm') = '" & Format(DTPMes.Value, "mm") & "' And To_Char(P.Fec_Prd,'yyyy') = '" & Year(DTPMes.Value) & "' Group By P.Linea, P.Batch")
           End If
           Set DataGridBatch2.DataSource = RProduccionLiberada
           
         MousePointer = 0
    'SALIDA
    ElseIf Index = 1 Then
        Unload Me
    End If
                      
End Sub

Private Sub Form_Load()
    DTPMes.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    RProduccion.Close
    RProduccionLiberada.Close
    Set RProduccion = Nothing
    Set RProduccionLiberada = Nothing
    If Err <> 0 Then
    End If
End Sub
