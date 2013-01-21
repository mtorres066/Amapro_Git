VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioInspeccion 
   Caption         =   "Inspeccion De Entradas A Inventario"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioInspeccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInspeccionar 
      Caption         =   "&Inspeccionar Todos"
      Height          =   975
      Left            =   9240
      Picture         =   "InventarioInspeccion.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DbGrid 
      Height          =   6855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column07 
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
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   975
      Left            =   10560
      Picture         =   "InventarioInspeccion.frx":39F4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdVerDatos 
      Caption         =   "&Ver Datos"
      Height          =   975
      Left            =   7920
      Picture         =   "InventarioInspeccion.frx":5A66
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtTra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numero Transaccion"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   480
      Width           =   1770
   End
End
Attribute VB_Name = "InventarioInspeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RVerDatos As New ADODB.Recordset

Private Sub CmdInspeccionar_Click()
        On Error Resume Next
        If Not IsNumeric(TxtTra.Text) Then
            MsgBox "Numero De Transaccion Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If

        Set RVerDatos = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtTra.Text & " And D.Estado = 'NO INSPECCIONADO' And D.FichaTecnica = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtTra.Text & " And UPPER(D.Estado) = 'NO INSPECCIONADO' And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
            End If
            
                If RVerDatos.RecordCount > 0 Then
                    Set DbGrid.DataSource = RVerDatos
                        Conexion.Execute "Update DetalleEntradasInventario Set Estado = 'INSPECCIONADO' Where Documento = " & TxtTra.Text & " And Estado = 'NO INSPECCIONADO'"
                Else
                    MsgBox "Esta Transaccion, No Tiene Bultos/Tarimas Pendientes De Inspeccionar", vbOKOnly + vbInformation, "Informacion"
                End If
            
                If Err <> 0 Then
                
                End If
                
                MsgBox "Transaccion Inpeccionada", vbOKOnly + vbInformation, "Informacion"

End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub CmdVerDatos_Click()
On Error Resume Next
        If Not IsNumeric(TxtTra.Text) Then
            MsgBox "Numero De Transaccion Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If

        Set RVerDatos = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtTra.Text & " And D.Estado = 'NO INSPECCIONADO' And D.FichaTecnica = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RVerDatos, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Bodega, D.Calidad, D.Estado, D.PesoEntrada, D.CantidadEntrada From DetalleEntradasInventario D, FichaTecnica F Where D.Documento = " & TxtTra.Text & " And UPPER(D.Estado) = 'NO INSPECCIONADO' And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
            End If
            
                If RVerDatos.RecordCount > 0 Then
                    Set DbGrid.DataSource = RVerDatos
                Else
                    MsgBox "Esta Transaccion, No Tiene Bultos/Tarimas Pendientes De Inspeccionar", vbOKOnly + vbInformation, "Informacion"
                End If
            
                If Err <> 0 Then
                
                End If

End Sub

Private Sub DbGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error Resume Next
        If ColIndex = 6 Then
            If DbGrid.Columns(6).Text = "A" Or DbGrid.Columns(6).Text = "I" Or DbGrid.Columns(6).Text = "R" Then
            Else
                Cancel = True
                MsgBox "Calidad Incorrecta"
            End If
        End If
        If ColIndex = 7 Then
            If DbGrid.Columns(7).Text = "INSPECCIONADO" Or DbGrid.Columns(7).Text = "NO INSPECCIONADO" Then
            Else
                Cancel = True
                MsgBox "Estado Incorrecto"
            End If
        End If
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If

End Sub

Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
                RVerDatos.Sort = RVerDatos.Fields(ColIndex).Name
End Sub

Private Sub TxtTra_GotFocus()
        TxtTra.SelStart = 0
        TxtTra.SelLength = Len(TxtTra.Text)
End Sub

Private Sub TxtTra_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
