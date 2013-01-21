VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioLiberacionTraslados 
   BackColor       =   &H00C00000&
   Caption         =   "Liberacion De Traslados De Inventario"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioLiberacionTraslados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid Dbgridrecepcion 
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
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
   Begin VB.TextBox TxtBod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox TxtReq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MskFec 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "InventarioLiberacionTraslados.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton CmdLiberar 
      Caption         =   "&Liberar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "InventarioLiberacionTraslados.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox TxtDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label LblBod 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega De Salida"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   11
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Requerido"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   2
      Left            =   2145
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaccion"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "InventarioLiberacionTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaTarima As New ADODB.Recordset
Dim REncabezadoEntradasProductoTerminado As New ADODB.Recordset
Dim RBuscaEncabezadoTraslados As New ADODB.Recordset
Dim RBuscaProducto As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaDetalleTraslados As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RDatos As New ADODB.Recordset

Dim VFechaEntrada As Date
Dim VFichaTecnica As String
Dim VTarima As Long
Dim VFechaProduccion As Date
Dim VLinea As String
Dim VBodegaATrasladar As String

Dim mensaje As String

Private Sub CmdLiberar_Click()
On Error Resume Next
                                    'BUSCA EL ESTADO DE EL DOCUMENTO
                                    Set RBuscaDocumento = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDocumento, "Select Estado From EncabezadoTrasladosInventario Where Documento = " & TxtDoc.Text)
                                            If RBuscaDocumento.RecordCount > 0 Then
                                                    If RBuscaDocumento!Estado = "LIBERADO" Then
                                                        MsgBox "Esta Transaccion Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
                                                        TxtDoc.SetFocus
                                                        Exit Sub
                                                    End If
                                            Else
                                                    MsgBox "Numero De Transaccion No Existe", vbOKOnly + vbExclamation, "Informacion"
                                                    TxtDoc.SetFocus
                                                    Exit Sub
                                            End If
                                    
                                    
                                    'PREGUNTA SI QUIERE LIBERAR
                                    mensaje = MsgBox("Está Seguro Liberar La Transaccion " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
                                    'SI DICE QUE NO SE SALE
                                    If mensaje = vbCancel Then
                                        Exit Sub
                                    End If
                                                    
                                    MousePointer = 11
                                                                        
                                    
                                    'BUSCA EL DETALLE DEL TRASLADO
                                    Set RBuscaDetalleTraslados = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDetalleTraslados, "Select FichaTecnica, Tarima, FechaProduccion, LineaProduccion, BodegaEntrada, CantidadReal From DetalleTrasladosInventario Where Documento = " & TxtDoc.Text)
                                    
                                    
                                            If RBuscaDetalleTraslados.RecordCount > 0 Then
                                                'INICIA LA TRANSACCION
                                                Conexion.BeginTrans
                                                
                                                    'CREA UN CICLO PARA RECORRER TODO EL TRASLADO
                                                    Do Until RBuscaDetalleTraslados.EOF
                                                    
                                                            VFichaTecnica = RBuscaDetalleTraslados!FichaTecnica
                                                            VTarima = RBuscaDetalleTraslados!Tarima
                                                            VFechaProduccion = RBuscaDetalleTraslados!FechaProduccion
                                                            VLinea = RBuscaDetalleTraslados!LineaProduccion
                                                    
                                                            'BUSCA TARIMA EN INVENTARIO (ENTRADAS PT) Y MODIFICA LA BODEGA
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Conexion.Execute ("update DetalleEntradasInventario set bodega = '" & RBuscaDetalleTraslados!BodegaEntrada & "', Saldo = " & RBuscaDetalleTraslados!CantidadReal & " Where FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                                                            Else
                                                                Conexion.Execute ("update DetalleEntradasInventario set bodega = '" & UCase(RBuscaDetalleTraslados!BodegaEntrada) & "', Saldo = " & RBuscaDetalleTraslados!CantidadReal & " Where UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "' And Tarima = " & VTarima & " And FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                                            End If
                                                            
                                                            If Err <> 0 Then
                                                                Conexion.RollbackTrans
                                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                                Err.Clear
                                                                Exit Sub
                                                            End If
                                                            
                                                        RBuscaDetalleTraslados.MoveNext
                                                    Loop
                                                        
                                                    'BUSCA EL ENCABEZADO DE LA ENTRADA PARA CAMBIAR EL ESTADO POR LIBERADO
                                                    Conexion.Execute "Update EncabezadoTrasladosInventario Set Estado = 'LIBERADO', Liberado = '" & GUsuario & "' Where Documento = " & TxtDoc.Text
                                                            
                                                            If Err <> 0 Then
                                                                Conexion.RollbackTrans
                                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                                Err.Clear
                                                            End If
                                                            
                                                'TERMINA LA TRANSACCION
                                                Conexion.CommitTrans
                                            End If
                                            
                                    MousePointer = 0
                                    
                                    MsgBox "Documento Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub TxtBod_Change()
            Set RBuscaBodega = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod.Text) & "'")
                End If
                If RBuscaBodega.RecordCount > 0 Then
                    LblBod.Caption = RBuscaBodega!Descripcion
                Else
                    LblBod.Caption = ""
                End If
End Sub

Private Sub TxtDoc_Change()
On Error Resume Next
    If IsNumeric(TxtDoc.Text) Then
            Set RDatos = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDatos, "Select DD.BodegaEntrada, DD.FechaProduccion, DD.LineaProduccion, DD.FichaTecnica, FT.Descrip, DD.Tarima, DD.CantidadSalida, DD.DiferenciaReqCorMas, DD.DiferenciaReqCor, DD.CantidadDesperdicio, DD.CantidadDesperdicioProveedor, DD.CantidadReal From DetalleTrasladosInventario DD, FichaTecnica FT Where DD.Documento = " & TxtDoc.Text & " And DD.FichaTecnica = FT.Esp_Tec")
                Else
                    Call Abrir_Recordset(RDatos, "Select DD.BodegaEntrada, DD.FichaTecnica, DD.Tarima, DD.FechaProduccion, DD.LineaProduccion, DD.CantidadSalida, DD.DifReqCorMas, DD.DifReqCor, DD.CantidadDesperdicio, DD.CantidadDesperdicioProveedor, DD.CantidadReal From DetalleTrasladosInventario DD, FichaTecnica FT Where DD.Documento = " & TxtDoc.Text & " And UPPER(DD.FichaTecnica) = UPPER(FT.Esp_Tec)")
                End If
                If RDatos.RecordCount > 0 Then
                
                End If
            Set Dbgridrecepcion.DataSource = RDatos
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE DESPACHOS
            Set RBuscaEncabezadoTraslados = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaEncabezadoTraslados, "Select Fecha, Requerido, Observaciones, BodegaSalida From EncabezadoTrasladosInventario Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoTraslados.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoTraslados!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoTraslados!fecha
                End If
                'REQUERIDO POR
                If IsNull(RBuscaEncabezadoTraslados!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = RBuscaEncabezadoTraslados!Requerido
                End If
                'OBSERVACIONES
                If IsNull(RBuscaEncabezadoTraslados!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RBuscaEncabezadoTraslados!Observaciones
                End If
                'BODEGA
                If IsNull(RBuscaEncabezadoTraslados!BodegaSalida) Then
                    TxtBod.Text = ""
                Else
                    TxtBod.Text = RBuscaEncabezadoTraslados!BodegaSalida
                End If
            Else
                MskFec.Text = ""
                TxtReq.Text = ""
                TxtObs.Text = ""
                TxtBod.Text = ""
            End If
    End If
      
End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Sub AnchoColumnas()
        Dbgridrecepcion.Columns(0).Width = "500"
        Dbgridrecepcion.Columns(1).Width = "1000"
        Dbgridrecepcion.Columns(2).Width = "500"
        Dbgridrecepcion.Columns(3).Width = "1300"
        Dbgridrecepcion.Columns(4).Width = "3000"
        Dbgridrecepcion.Columns(5).Width = "500"
        Dbgridrecepcion.Columns(6).Width = "700"
        Dbgridrecepcion.Columns(7).Width = "700"
        Dbgridrecepcion.Columns(8).Width = "700"
        Dbgridrecepcion.Columns(9).Width = "700"
        Dbgridrecepcion.Columns(10).Width = "700"
        Dbgridrecepcion.Columns(11).Width = "700"
End Sub
