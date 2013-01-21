VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form OCompraMP 
   BorderStyle     =   0  'None
   Caption         =   "Ordenes de Compra de MP Pendientes"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame_Compras_MP 
      Caption         =   "Ordenes de Compra de Materia Prima"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Cmd_Cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9480
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton Cmd_Salir 
         BackColor       =   &H80000002&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
      Begin MSDataGridLib.DataGrid Dg_Compras_MP 
         Height          =   6615
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11668
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
               LCID            =   2058
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
               LCID            =   2058
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
End
Attribute VB_Name = "OCompraMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaOrdenCompraMP As New ADODB.Recordset

Private Sub Cmd_Cancelar_Click()
    
    OCompraMP.Visible = False
    
End Sub

Private Sub Cmd_Salir_Click()
    
    OCompraMP.Visible = False
    
End Sub

'CSEH: ErrResumeNext
Private Sub Form_Load()
    
    On Error Resume Next
            
        Frame_Compras_MP.Visible = True
                  
    ' ABRIMOS RECORDSET PARA SACAR LAS O.C. DE MP DE DICHO PRODUCTO CAPTURADO
    Set RBuscaOrdenCompraMP = New ADODB.Recordset
    If GOrigenDeDatos = "AmaproAccess" Then
        Call Abrir_Recordset(RBuscaOrdenCompraMP, "SELECT DISTINCT " & _
                            "E.Fecha, E.Documento, E.Proveedor, " & _
                            "P.Descripcion, D.Codigo, D.CantidadPedido, " & _
                            "D.SaldoPorEntregar " & _
                            "FROM " & _
                            "EncabezadoPedidosProveedores E, " & _
                            "DetallePedidosProveedores D, " & _
                            "Proveedores P " & _
                            "WHERE " & _
                            "E.Documento = D.Documento AND " & _
                            "D.SaldoPorEntregar > 0 AND " & _
                            "E.Proveedor = P.CodigoProveedor AND " & _
                            "D.Codigo LIKE '642-E-506'")
    Else 'ORACLE

    End If
        
        Set Dg_Compras_MP.DataSource = RBuscaOrdenCompraMP
    
End Sub
