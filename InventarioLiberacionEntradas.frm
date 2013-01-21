VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioLiberacionEntradas 
   BackColor       =   &H00C00000&
   Caption         =   "Liberacion De Entradas De Inventario"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioLiberacionEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "InventarioLiberacionEntradas.frx":030A
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
      Picture         =   "InventarioLiberacionEntradas.frx":237C
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
   Begin MSDataGridLib.DataGrid DbGridRecepcion 
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "FechaProduccion"
         Caption         =   "Fecha"
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
         DataField       =   "Linea"
         Caption         =   "Linea"
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
      BeginProperty Column02 
         DataField       =   "FichaTecnica"
         Caption         =   "Ficha Tecnica"
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
      BeginProperty Column03 
         DataField       =   "Descrip"
         Caption         =   "Descripcion"
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
      BeginProperty Column04 
         DataField       =   "Tarima"
         Caption         =   "Tarima"
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
      BeginProperty Column05 
         DataField       =   "Bodega"
         Caption         =   "Bodega"
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
      BeginProperty Column06 
         DataField       =   "Calidad"
         Caption         =   "Calidad"
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
      BeginProperty Column07 
         DataField       =   "Estado"
         Caption         =   "Estado"
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
      BeginProperty Column08 
         DataField       =   "PesoEntrada"
         Caption         =   "PesoEntrada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4106
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "CantidadEntrada"
         Caption         =   "Cantidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4106
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   3465.071
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
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
Attribute VB_Name = "InventarioLiberacionEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RDetalleEntradasProductoTerminado As New ADODB.Recordset
Dim REncabezadoEntradasProductoTerminado As New ADODB.Recordset
Dim RBuscaEncabezadoEntradas As New ADODB.Recordset
Dim RBuscaCorrelativo As New ADODB.Recordset
Dim RBuscaPedido As New ADODB.Recordset
Dim RBuscaProducto As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RTarimas As New ADODB.Recordset
Dim RVerDatos As New ADODB.Recordset

Dim VFechaEntrada As Date
Dim VNumeroPedido As Double
Dim VCantidadEntrada As Double
Dim VBodega As String
Dim VProducto As String
Dim VDiasDeAtraso As Long
Dim VCorrelativo As Double

Dim mensaje As String

Private Sub CmdLiberar_Click()
On Error Resume Next
                                    'VERIFICA SI ES NUMERICO
                                    If Not IsNumeric(TxtDoc.Text) Then
                                        MsgBox "El Numero De Documento Debe Ser Numerico", vbOKOnly + vbExclamation, "Verificacion"
                                        TxtDoc.SetFocus
                                        Exit Sub
                                    End If
                                    
                                    'BUSCA EL ESTADO DE EL DOCUMENTO
                                    Set RBuscaDocumento = New ADODB.Recordset
                                        Call Abrir_Recordset(RBuscaDocumento, "Select Estado From EncabezadoEntradasInventario Where Documento = " & TxtDoc.Text)
                                                If RBuscaDocumento.RecordCount > 0 Then
                                                        If RBuscaDocumento!Estado = "LIBERADO" Then
                                                            MsgBox "Esta Documento Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                                                            TxtDoc.SetFocus
                                                            Exit Sub
                                                        End If
                                                Else
                                                        MsgBox "Numero De Transaccion No Existe", vbOKOnly + vbExclamation, "Informacion"
                                                        TxtDoc.SetFocus
                                                        Exit Sub
                                                End If
                                                
                                    Set RVerDatos = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtDoc.Text & " And D.Estado = 'NO INSPECCIONADO' And D.FichaTecnica = F.Esp_Tec")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtDoc.Text & " And UPPER(D.Estado) = 'NO INSPECCIONADO' And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                                        End If
                                        
                                            If RVerDatos.RecordCount > 0 Then
                                                MsgBox "Todavia Hay Bultos/Tarimas Pendientes De Inspeccionar", vbOKOnly + vbInformation, "Informacion"
                                                Exit Sub
                                            Else
                                                
                                            End If
                                    
                                    
                                    
                                    'PREGUNTA SI QUIERE LIBERAR
                                    mensaje = MsgBox("Está Seguro Liberar La Transaccion " & TxtDoc.Text, vbOKCancel + vbCritical + vbDefaultButton2, "Verificacion")
                                    
                                    'SI DICE QUE NO SE SALE
                                    If mensaje = vbCancel Then
                                        Exit Sub
                                    End If
                                                    
                                    MousePointer = 11
                                    
                                    'BUSCA LA FECHA DE Documento
                                    Set RBuscaDocumento = New ADODB.Recordset
                                    Call Abrir_Recordset(RBuscaDocumento, "Select FechaEntrada From EncabezadoEntradasInventario Where Documento = " & TxtDoc.Text)
                                    If RBuscaDocumento.RecordCount > 0 Then
                                        VFechaEntrada = RBuscaDocumento!FechaEntrada
                                    Else
                                        VFechaEntrada = Date
                                    End If
                                                                                                                    
                                    'BUSCA EL ENCABEZADO DE LA ENTRADA PARA CAMBIAR EL ESTADO POR LIBERADO
                                    Conexion.Execute "Update EncabezadoEntradasInventario Set Estado = 'LIBERADO', Liberado = '" & GUsuario & "' Where Documento = " & TxtDoc.Text
                                            
                                    MousePointer = 0
                                    
                                    MsgBox "Documento Liberado Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub TxtDoc_Change()
On Error Resume Next
      If IsNumeric(TxtDoc.Text) Then
                       Set RTarimas = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtDoc.Text & " And  D.FichaTecnica = F.Esp_Tec")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtDoc.Text & " And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                                        End If
                                            If RTarimas.RecordCount > 0 Then
                                                Set Dbgridrecepcion.DataSource = RTarimas
                                            Else
                                                
                                            End If
                                    
      
                    
                    
            
            'BUSCA EL ENCABEZADO DE ENTRADAS
            Set RBuscaEncabezadoEntradas = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaEncabezadoEntradas, "Select FechaEntrada, Requerido, Observaciones From EncabezadoEntradasInventario Where Documento = " & TxtDoc.Text)
            If RBuscaEncabezadoEntradas.RecordCount > 0 Then
                'FECHA DE TRASLADO
                If IsNull(RBuscaEncabezadoEntradas!FechaEntrada) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = RBuscaEncabezadoEntradas!FechaEntrada
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
      Else
             If IsNumeric(TxtDoc.Text) Then
                Set RTarimas = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select DE.Bodega, DE.FichaTecnica, FT.Descrip, DE.Cantidad, DE.Tarima, DE.FechaProduccion, DE.Linea, DE.Calidad, DE.Batch From DetalleEntradasInventario DE, FichaTecnica FT Where DE.FichaTecnica = FT.Esp_Tec And DE.Documento = " & TxtDoc.Text & " Order By DE.FechaProduccion, DE.Linea, DE.Tarima")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select DE.Bodega, DE.FichaTecnica, FT.Descrip, DE.Cantidad, DE.Tarima, DE.FechaProduccion, DE.Linea, DE.Calidad, DE.Batch From DetalleEntradasInventario DE, FichaTecnica FT Where UPPER(DE.FichaTecnica) = UPPER(FT.Esp_Tec) And DE.Documento = " & TxtDoc.Text & " Order By DE.FechaProduccion, DE.Linea, DE.Tarima")
                    End If
                Set Dbgridrecepcion.DataSource = RTarimas
                
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

