VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioLiberacionSalidas 
   BackColor       =   &H00C00000&
   Caption         =   "Liberacion De Salidas De Inventario"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioLiberacionSalidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DbGridRecepcion 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
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
      Picture         =   "InventarioLiberacionSalidas.frx":030A
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
      Picture         =   "InventarioLiberacionSalidas.frx":237C
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "InventarioLiberacionSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaTarima As New ADODB.Recordset
Dim REncabezadoEntradasProductoTerminado As New ADODB.Recordset
Dim RBuscaEncabezadoEntradas As New ADODB.Recordset
Dim RBuscaProducto As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaDetalleSalidas As New ADODB.Recordset
Dim RDatos As New ADODB.Recordset

Dim VFechaEntrada As Date
Dim VFichaTecnica As String
Dim VTarima As Long
Dim VFechaProduccion As Date
Dim VLinea As String

Dim mensaje As String

Private Sub CmdLiberar_Click()
On Error Resume Next
                                    
                                    'BUSCA EL ESTADO DE EL DOCUMENTO
                                    Set RBuscaDocumento = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDocumento, "Select Estado From EncabezadoSalidasInventario Where Documento = " & TxtDoc.Text)
                                            If RBuscaDocumento.RecordCount > 0 Then
                                                    If RBuscaDocumento!Estado = "LIBERADO" Then
                                                        MsgBox "Esta Documento Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                                                        TxtDoc.SetFocus
                                                        Exit Sub
                                                    End If
                                            Else
                                                    MsgBox "Numero De Documento No Existe", vbOKOnly + vbExclamation, "Informacion"
                                                    TxtDoc.SetFocus
                                                    Exit Sub
                                            End If
                                            
                                    
                                    'PREGUNTA SI QUIERE LIBERAR
                                    mensaje = MsgBox("Está Seguro Liberar El Documento " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
                                    'SI DICE QUE NO SE SALE
                                    If mensaje = vbCancel Then
                                        Exit Sub
                                    End If
                                                    
                                    MousePointer = 11
                                    
                                    
                                    'BUSCA LA FECHA DE Documento
                                    Set RBuscaDocumento = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDocumento, "Select Fecha From EncabezadoSalidasInventario Where Documento = " & TxtDoc.Text)
                                            If RBuscaDocumento.RecordCount > 0 Then
                                                VFechaEntrada = RBuscaDocumento!fecha
                                            Else
                                                VFechaEntrada = Date
                                            End If
                                                                                                                            
                                    'BUSCA EL DETALLE DE DESPACHOS CON ESTE DOCUMENTO
                                    Set RBuscaDetalleSalidas = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDetalleSalidas, "Select FichaTecnica, Tarima, FechaProduccion, Linea, Cantidad From DetalleSalidasInventario Where Documento = " & TxtDoc.Text)
                                    
                                        'CREA UN CICLO PARA DESCARGAR TARIMA POR TARIMA DEL INVENTARIO
                                        Do Until RBuscaDetalleSalidas.EOF
                                        
                                                VFichaTecnica = RBuscaDetalleSalidas!FichaTecnica
                                                VTarima = RBuscaDetalleSalidas!Tarima
                                                VFechaProduccion = RBuscaDetalleSalidas!Fechaproduccion
                                                VLinea = RBuscaDetalleSalidas!Linea
                                        
                                                'BUSCA TARIMA
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & RBuscaDetalleSalidas!Cantidad & " Where FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'"
                                                Else 'ORACLE
                                                    Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & RBuscaDetalleSalidas!Cantidad & " Where UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "' And Tarima = " & VTarima & " And FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'"
                                                End If
                                                
                                            RBuscaDetalleSalidas.MoveNext
                                        Loop
                                            
                                            
                                    'BUSCA EL ENCABEZADO DE LA ENTRADA PARA CAMBIAR EL ESTADO POR LIBERADO
                                    Conexion.Execute ("Update EncabezadoSalidasInventario Set Estado = 'LIBERADO', Liberado = '" & GUsuario & "' Where Documento = " & TxtDoc.Text)
                                            
                                            
                                    MousePointer = 0
                                    
                                    MsgBox "Documento Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub TxtDoc_Change()
On Error Resume Next
      If IsNumeric(TxtDoc.Text) Then
            Set RDatos = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDatos, "Select DD.Bodega, DD.FichaTecnica, FT.Descrip, DD.Cantidad, DD.Tarima, DD.FechaProduccion, DD.Linea, DD.Calidad, DD.Batch From DetalleSalidasInventario DD, FichaTecnica FT Where DD.FichaTecnica = FT.Esp_Tec And DD.Documento = " & TxtDoc.Text & " Order By DD.FechaProduccion, DD.Linea, DD.Tarima")
                Else
                    Call Abrir_Recordset(RDatos, "Select DD.Bodega, DD.FichaTecnica, FT.Descrip, DD.Cantidad, DD.Tarima, DD.FechaProduccion, DD.Linea, DD.Calidad, DD.Batch From DetalleSalidasInventario DD, FichaTecnica FT Where UPPER(DD.FichaTecnica) = UPPER(FT.Esp_Tec) And DD.Documento = " & TxtDoc.Text & " Order By DD.FechaProduccion, DD.Linea, DD.Tarima")
                End If
            
            Set DbGridRecepcion.DataSource = RDatos
            AnchoColumnas
            
            'BUSCA EL ENCABEZADO DE DESPACHOS
            Set RBuscaEncabezadoEntradas = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaEncabezadoEntradas, "Select Fecha, Requerido, Observaciones From EncabezadoSalidasInventario Where Documento = " & TxtDoc.Text)
                    If RBuscaEncabezadoEntradas.RecordCount > 0 Then
                        'FECHA DE TRASLADO
                        If IsNull(RBuscaEncabezadoEntradas!fecha) Then
                            MskFec.Text = ""
                        Else
                            MskFec.Text = RBuscaEncabezadoEntradas!fecha
                        End If
                        'REQUERIDO POR
                        If IsNull(RBuscaEncabezadoEntradas!Requerido) Then
                            TxtReq.Text = ""
                        Else
                            TxtReq.Text = RBuscaEncabezadoEntradas!Requerido
                        End If
                        'OBSERVACIONES
                        If IsNull(RBuscaEncabezadoEntradas!Observaciones) Then
                            TxtObs.Text = ""
                        Else
                            TxtObs.Text = RBuscaEncabezadoEntradas!Observaciones
                        End If
                    Else
                        MskFec.Text = ""
                        TxtReq.Text = ""
                        TxtObs.Text = ""
                    End If
        End If
      
End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Sub AnchoColumnas()
        DbGridRecepcion.Columns(0).Width = "500"
        DbGridRecepcion.Columns(1).Width = "1400"
        DbGridRecepcion.Columns(2).Width = "4000"
        DbGridRecepcion.Columns(3).Width = "700"
        DbGridRecepcion.Columns(4).Width = "500"
        DbGridRecepcion.Columns(5).Width = "1000"
        DbGridRecepcion.Columns(6).Width = "500"
        DbGridRecepcion.Columns(7).Width = "500"
End Sub
