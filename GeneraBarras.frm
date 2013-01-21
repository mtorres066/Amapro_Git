VERSION 5.00
Begin VB.Form GeneraBarras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador De Barras"
   ClientHeight    =   4770
   ClientLeft      =   2700
   ClientTop       =   1380
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "GeneraBarras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Formatos De Impresion"
      Height          =   1455
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton OptFor 
         Caption         =   "Boleta De Caja"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton OptFor 
         Caption         =   "Cedula"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton OptFor 
         Caption         =   "Boleta Identificacion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtLinea 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtBatch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton OptOpcion 
      Caption         =   "Por Texto"
      ForeColor       =   &H00000000&
      Height          =   192
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   1095
      Left            =   7560
      Picture         =   "GeneraBarras.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton OptOpcion 
      Caption         =   "Por Batch De Inventario"
      ForeColor       =   &H00FF0000&
      Height          =   192
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.OptionButton OptOpcion 
      Caption         =   "Por Batch De Produccion Liberada"
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.OptionButton OptOpcion 
      Caption         =   "Por Batch De Produccion"
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Salida"
      Height          =   1095
      Left            =   8760
      Picture         =   "GeneraBarras.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   120
      TabIndex        =   13
      Text            =   "25-05-2004-02-200308-64-23"
      Top             =   2880
      Width           =   7455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   120
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   639
      TabIndex        =   12
      Top             =   3480
      Width           =   9615
   End
   Begin VB.Label LblLinea 
      Alignment       =   1  'Right Justify
      Caption         =   "Linea"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblBatch 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese El Texto Para El Codigo De Barra"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "GeneraBarras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RBatch As New ADODB.Recordset
Dim VContador As Long


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

MousePointer = 11
        'BOLETA DE IDENTIFICACION
        If OptFor.Item(0).Value = True Then
            'VContador = 13000
            VContador = 600
        'CEDULA
        ElseIf OptFor.Item(1).Value = True Then
            'VContador = 7200
            VContador = 600
        'BOLETA CAJA
        ElseIf OptFor.Item(2).Value = True Then
            'VContador = 10000
            VContador = 600
        End If
                
        Set RBatch = New ADODB.Recordset
        'PRODUCCION
        If OptOpcion.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBatch, "Select Barra From Produccion Where Batch = " & TxtBatch & " And Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBatch, "Select Barra From Produccion Where Batch = " & TxtBatch & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
        'PRODUCCION LIBERADA
        ElseIf OptOpcion.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBatch, "Select Barra From ProduccionLiberada Where Batch = " & TxtBatch & " And Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBatch, "Select Barra From ProduccionLiberada Where Batch = " & TxtBatch & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
        'INVENTARIO
        ElseIf OptOpcion.Item(2).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBatch, "Select Barra From DetalleEntradasInventario Where Batch = " & TxtBatch & " And Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBatch, "Select Barra From DetalleEntradasInventario Where Batch = " & TxtBatch & " And UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
        'TEXTO
        ElseIf OptOpcion.Item(4).Value = True Then
                'BOLETA DE IDENTIFACION
                If OptFor.Item(0).Value = True Then
                    'Printer.PaintPicture Picture1, 3000, 13000
                    Printer.PaintPicture Picture1, 3000, 600
                'CEDULA
                ElseIf OptFor.Item(1).Value = True Then
                    'Printer.PaintPicture Picture1, 7000, 7200
                    Printer.PaintPicture Picture1, 3000, 600
                'BOLETA CAJA
                ElseIf OptFor.Item(2).Value = True Then
                    'Printer.PaintPicture Picture1, 500, 10000
                    Printer.PaintPicture Picture1, 3000, 600
                End If
                
                Printer.EndDoc
        End If
        
                'SI ELIGE POR TEXTO NO HACE NADA
                If OptOpcion.Item(4).Value = True Then
                Else
                    If RBatch.RecordCount > 0 Then
                            Do Until RBatch.EOF
                                    If IsNull(RBatch!Barra) Then
                                    Else
                                        'CONVIERTE EL CODIGO DE BARRAS
                                        Call DrawBarcode(RBatch!Barra, Picture1)
                                    
                                    
                                        MinWidth = 32 * Text1.Left + Text1.Width
                                        pw = 32 * Picture1.Left + Picture1.Width
                                        fw = MinWidth
                                        If pw > fw Then fw = pw
                                        
                                        'IMPRIME LA BARRA
                                        'BOLETA DE IDENTIFACION
                                        If OptFor.Item(0).Value = True Then
                                            'Printer.PaintPicture Picture1, 3000, VContador
                                            Printer.PaintPicture Picture1, 3000, VContador
                                        'CEDULA
                                        ElseIf OptFor.Item(1).Value = True Then
                                            'Printer.PaintPicture Picture1, 7000, VContador
                                            Printer.PaintPicture Picture1, 3000, VContador
                                        'BOLETA BLANCA
                                        ElseIf OptFor.Item(2).Value = True Then
                                            'Printer.PaintPicture Picture1, 500, VContador
                                            Printer.PaintPicture Picture1, 3000, VContador
                                        End If
                                        
                                        'Printer.EndDoc
                                        'LE SUMA 2000 PARA QUE NO CAIGA EN LA MISMA POSICION
                                        VContador = VContador + 2000
                                    End If
                                    
                                RBatch.MoveNext
                            Loop
                    End If
                                    'PARA QUE EMPIEZE A IMPRIMIR
                                    Printer.EndDoc
                End If
                
MousePointer = 0
                MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"

End Sub

Private Sub Form_Activate()
        Picture1.ScaleMode = 3
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
        
        Call DrawBarcode(Text1, Picture1)

End Sub

Private Sub OptOpcion_Click(Index As Integer)
        'PRODUCCION, PRODUCCION LIBERADA, Y PRODUCTO TERMINADO
        If Index = 0 Or Index = 1 Or Index = 2 Then
            Label1.Visible = False
            Text1.Visible = False
            Picture1.Visible = False
            
            TxtBatch.Visible = True
            TxtLinea.Visible = True
            
            LblBatch.Visible = True
            LblBatch.Caption = "Batch"
            TxtBatch.SetFocus
            LblLinea.Visible = True
            
        End If
        
        'MATERIA PRIMA
        If Index = 3 Then
            Label1.Visible = False
            Text1.Visible = False
            Picture1.Visible = False
            
            TxtBatch.Visible = True
            TxtLinea.Visible = False
            
            LblBatch.Visible = True
            LblBatch.Caption = "Transaccion"
            TxtBatch.SetFocus
            LblLinea.Visible = False
        End If
        'SI ELIGE POR TEXTO
        If Index = 4 Then
            Label1.Visible = True
            Text1.Visible = True
            Text1.SetFocus
            Picture1.Visible = True
            TxtBatch.Visible = False
            TxtLinea.Visible = False
            LblBatch.Visible = False
            LblLinea.Visible = False
        End If
        
End Sub

Private Sub Text1_Change()
    
    Call DrawBarcode(Text1, Picture1)
    
    MinWidth = 32 * Text1.Left + Text1.Width
    pw = 32 * Picture1.Left + Picture1.Width
    fw = MinWidth
    If pw > fw Then fw = pw
    'Form1.Width = fw

End Sub

Private Sub Text1_GotFocus()
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
End Sub
