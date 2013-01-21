VERSION 5.00
Begin VB.Form PasarDatosAccessAOracle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasar Datos De Access A Oracle"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opt 
      Caption         =   "Diez"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Nueve"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Ocho"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Tres 2"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Seis"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Cinco"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Cuatro"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Tres"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Dos"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Uno"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar Transferencia"
      Height          =   1575
      Left            =   3960
      Picture         =   "PasarDatosAccessAOracle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
   End
End
Attribute VB_Name = "PasarDatosAccessAOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RCambiaDocumento As New ADODB.Recordset
Dim RAtributos As New ADODB.Recordset
Dim RAjustesMP As New ADODB.Recordset
Dim RAjustesPT As New ADODB.Recordset
Dim RBatch As New ADODB.Recordset
Dim RBatchDatos As New ADODB.Recordset
Dim RBodegasMateriaPrima As New ADODB.Recordset
Dim RBodegasMateriaPrimaGrupos As New ADODB.Recordset
Dim RBodegasProductoTerminado As New ADODB.Recordset
Dim RBodegasProductoTerminadoGrupos As New ADODB.Recordset
Dim RCapturaDesperdicioMP As New ADODB.Recordset
Dim RCapturaRutinas As New ADODB.Recordset
Dim RCierreTarimas As New ADODB.Recordset
Dim RClientes As New ADODB.Recordset
Dim RCobrosProveedor As New ADODB.Recordset
Dim RCorrelativosMP As New ADODB.Recordset
Dim RDefectos As New ADODB.Recordset
Dim RDefectosMP As New ADODB.Recordset
Dim RDesperdicioPacas As New ADODB.Recordset
Dim RDetalleCapturaParos As New ADODB.Recordset
Dim RDetalleCierrePedidosClientes As New ADODB.Recordset
Dim RDetalleCierrePedidosProveedores As New ADODB.Recordset
Dim RDetalleConsumoMP As New ADODB.Recordset
Dim RDetalleDespachosPT As New ADODB.Recordset
Dim RDetalleEgresosMP As New ADODB.Recordset
Dim RDetalleEntradasMP As New ADODB.Recordset
Dim RDetalleEntradasPT As New ADODB.Recordset
Dim RDetalleOrdenProduccion As New ADODB.Recordset
Dim RDetallePedidosClientes As New ADODB.Recordset
Dim RDetallePedidosProveedores As New ADODB.Recordset
Dim RDetalleProduccionOrdenLiberada As New ADODB.Recordset
Dim RDetalleProduccionPorOrden As New ADODB.Recordset
Dim RDetalleTrasladosMP As New ADODB.Recordset
Dim RDetalleTrasladosPT As New ADODB.Recordset
Dim RDocumentos As New ADODB.Recordset
Dim REmpleados As New ADODB.Recordset
Dim REmpleadosCapturaAumentos As New ADODB.Recordset
Dim REmpleadosCapturaCursos As New ADODB.Recordset
Dim REmpleadosCapturaFaltas As New ADODB.Recordset
Dim REmpleadosCursos As New ADODB.Recordset
Dim REmpleadosDepartamentos As New ADODB.Recordset
Dim REmpleadosEscolaridad As New ADODB.Recordset
Dim REmpleadosFactores As New ADODB.Recordset
Dim REmpleadosFaltas As New ADODB.Recordset
Dim REmpleadosGrupos As New ADODB.Recordset
Dim REmpleadosHabilidades As New ADODB.Recordset
Dim REmpleadosHabEmp As New ADODB.Recordset
Dim REmpleadosHabilidadesPuestos As New ADODB.Recordset
Dim REmpleadosHijos As New ADODB.Recordset
Dim REmpleadosPuestos As New ADODB.Recordset
Dim REncabezadoCapturaParos As New ADODB.Recordset
Dim REncabezadoCierrePedidosClientes As New ADODB.Recordset
Dim REncabezadoCierrePedidosProveedores As New ADODB.Recordset
Dim REncabezadoDespachosPT As New ADODB.Recordset
Dim REncabezadoEgresosMP As New ADODB.Recordset
Dim REncabezadoEntradasMP As New ADODB.Recordset
Dim REncabezadoEntradasPT As New ADODB.Recordset
Dim REncabezadoOrdenProduccion As New ADODB.Recordset
Dim REncabezadoPedidosClientes As New ADODB.Recordset
Dim REncabezadoPedidosProveedores As New ADODB.Recordset
Dim REncabezadoTrasladosMP As New ADODB.Recordset
Dim REncabezadoTrasladosPT As New ADODB.Recordset
Dim RFichaTecnica As New ADODB.Recordset
Dim RFichaTecnicaConMateriaPrima As New ADODB.Recordset
Dim RFichaTecnicaTipos As New ADODB.Recordset
Dim RGraficaDeParos As New ADODB.Recordset
Dim RHistograma As New ADODB.Recordset
Dim RInventario As New ADODB.Recordset
Dim RLineas As New ADODB.Recordset
Dim RLineasBultos As New ADODB.Recordset
Dim RNumerosIngresosProcesados As New ADODB.Recordset
Dim RParos As New ADODB.Recordset
Dim RParosGrupos As New ADODB.Recordset
Dim RPasadas As New ADODB.Recordset
Dim RPedidosProveedoresPorcentajeNoConforme As New ADODB.Recordset
Dim RProcesosMP As New ADODB.Recordset
Dim RProduccion As New ADODB.Recordset
Dim RProduccionConDefectos As New ADODB.Recordset
Dim RProduccionConMP As New ADODB.Recordset
Dim RProduccionLiberada As New ADODB.Recordset
Dim RProduccionLiberadaConDefectos As New ADODB.Recordset
Dim RProduccionLiberadaConTarimas As New ADODB.Recordset
Dim RProveedores As New ADODB.Recordset
Dim RProveedoresGrupos As New ADODB.Recordset
Dim RRutinas As New ADODB.Recordset
Dim RTiposMP As New ADODB.Recordset
Dim RTiposEntradasMP As New ADODB.Recordset
Dim RTransportistas As New ADODB.Recordset
Dim RTurnos As New ADODB.Recordset
Dim RUnidadMedida As New ADODB.Recordset
Dim RUsuarios As New ADODB.Recordset
Dim RVariablesDescripcion As New ADODB.Recordset
Dim RVariablesMedia As New ADODB.Recordset
Dim RVentas As New ADODB.Recordset
Dim RVentasDetalle As New ADODB.Recordset
Dim RBuscaFechaEntrada As New ADODB.Recordset

'VARIABLES PARA ORACLE
Dim RAtributos2 As New ADODB.Recordset
Dim RAjustesMP2 As New ADODB.Recordset
Dim RAjustesPT2 As New ADODB.Recordset
Dim RBatch2 As New ADODB.Recordset
Dim RBatchDatos2 As New ADODB.Recordset
Dim RBodegasMateriaPrima2 As New ADODB.Recordset
Dim RBodegasMateriaPrimaGrupos2 As New ADODB.Recordset
Dim RBodegasProductoTerminado2 As New ADODB.Recordset
Dim RBodegasProductoTerminadoGrupos2 As New ADODB.Recordset
Dim RCapturaDesperdicioMP2 As New ADODB.Recordset
Dim RCapturaRutinas2 As New ADODB.Recordset
Dim RCierreTarimas2 As New ADODB.Recordset
Dim RClientes2 As New ADODB.Recordset
Dim RCobrosProveedor2 As New ADODB.Recordset
Dim RCorrelativosMP2 As New ADODB.Recordset
Dim RDefectos2 As New ADODB.Recordset
Dim RDefectosMP2 As New ADODB.Recordset
Dim RDesperdicioPacas2 As New ADODB.Recordset
Dim RDetalleCapturaParos2 As New ADODB.Recordset
Dim RDetalleCierrePedidosClientes2 As New ADODB.Recordset
Dim RDetalleCierrePedidosProveedores2 As New ADODB.Recordset
Dim RDetalleConsumoMP2 As New ADODB.Recordset
Dim RDetalleDespachosPT2 As New ADODB.Recordset
Dim RDetalleEgresosMP2 As New ADODB.Recordset
Dim RDetalleEntradasMP2 As New ADODB.Recordset
Dim RDetalleEntradasPT2 As New ADODB.Recordset
Dim RDetalleOrdenProduccion2 As New ADODB.Recordset
Dim RDetallePedidosClientes2 As New ADODB.Recordset
Dim RDetallePedidosProveedores2 As New ADODB.Recordset
Dim RDetalleProduccionOrdenLiberada2 As New ADODB.Recordset
Dim RDetalleProduccionPorOrden2 As New ADODB.Recordset
Dim RDetalleTrasladosPT2 As New ADODB.Recordset
Dim RDetalleTrasladosMP2 As New ADODB.Recordset
Dim RDocumentos2 As New ADODB.Recordset
Dim REmpleados2 As New ADODB.Recordset
Dim REmpleadosCapturaAumentos2 As New ADODB.Recordset
Dim REmpleadosCapturaCursos2 As New ADODB.Recordset
Dim REmpleadosCapturaFaltas2 As New ADODB.Recordset
Dim REmpleadosCursos2 As New ADODB.Recordset
Dim REmpleadosDepartamentos2 As New ADODB.Recordset
Dim REmpleadosEscolaridad2 As New ADODB.Recordset
Dim REmpleadosFactores2 As New ADODB.Recordset
Dim REmpleadosFaltas2 As New ADODB.Recordset
Dim REmpleadosGrupos2 As New ADODB.Recordset
Dim REmpleadosHabilidades2 As New ADODB.Recordset
Dim REmpleadosHabEmp2 As New ADODB.Recordset
Dim REmpleadosHabilidadesPuestos2 As New ADODB.Recordset
Dim REmpleadosHijos2 As New ADODB.Recordset
Dim REmpleadosPuestos2 As New ADODB.Recordset
Dim REncabezadoCapturaParos2 As New ADODB.Recordset
Dim REncabezadoCierrePedidosClientes2 As New ADODB.Recordset
Dim REncabezadoCierrePedidosProveedores2 As New ADODB.Recordset
Dim REncabezadoDespachosPT2 As New ADODB.Recordset
Dim REncabezadoEgresosMP2 As New ADODB.Recordset
Dim REncabezadoEntradasMP2 As New ADODB.Recordset
Dim REncabezadoEntradasPT2 As New ADODB.Recordset
Dim REncabezadoOrdenProduccion2 As New ADODB.Recordset
Dim REncabezadoPedidosClientes2 As New ADODB.Recordset
Dim REncabezadoPedidosProveedores2 As New ADODB.Recordset
Dim REncabezadoTrasladosMP2 As New ADODB.Recordset
Dim REncabezadoTrasladosPT2 As New ADODB.Recordset
Dim RFichaTecnica2 As New ADODB.Recordset
Dim RFichaTecnicaConMateriaPrima2 As New ADODB.Recordset
Dim RFichaTecnicaTipos2 As New ADODB.Recordset
Dim RGraficaDeParos2 As New ADODB.Recordset
Dim RHistograma2 As New ADODB.Recordset
Dim RInventario2 As New ADODB.Recordset
Dim RLineas2 As New ADODB.Recordset
Dim RLineasBultos2 As New ADODB.Recordset
Dim RNumerosIngresosProcesados2 As New ADODB.Recordset
Dim RParos2 As New ADODB.Recordset
Dim RParosGrupos2 As New ADODB.Recordset
Dim RPasadas2 As New ADODB.Recordset
Dim RPedidosProveedoresPorcentajeNoConforme2 As New ADODB.Recordset
Dim RProcesosMP2 As New ADODB.Recordset
Dim RProduccion2 As New ADODB.Recordset
Dim RProduccionConDefectos2 As New ADODB.Recordset
Dim RProduccionConMP2 As New ADODB.Recordset
Dim RProduccionLiberada2 As New ADODB.Recordset
Dim RProduccionLiberadaConDefectos2 As New ADODB.Recordset
Dim RProduccionLiberadaConTarimas2 As New ADODB.Recordset
Dim RProveedores2 As New ADODB.Recordset
Dim RProveedoresGrupos2 As New ADODB.Recordset
Dim RRutinas2 As New ADODB.Recordset
Dim RTiposMP2 As New ADODB.Recordset
Dim RTiposEntradasMP2 As New ADODB.Recordset
Dim RTransportistas2 As New ADODB.Recordset
Dim RTurnos2 As New ADODB.Recordset
Dim RUnidadMedida2 As New ADODB.Recordset
Dim Rusuarios2 As New ADODB.Recordset
Dim RVariablesDescripcion2 As New ADODB.Recordset
Dim RVariablesMedia2 As New ADODB.Recordset
Dim RVentas2 As New ADODB.Recordset
Dim RVentasDetalle2 As New ADODB.Recordset

Dim Cont As Integer
Dim VProduccionInterna As Integer
Dim VProduccionLiberada As Integer
        
Private Sub Command1_Click()
On Error Resume Next
        MousePointer = 11
            'PARA CUANDO EMPIEZA EL PROCESO
            If Opt.Item(0).Value = True Then
                Text1.Text = "Uno " & Text1.Text & Time & vbCrLf
                uno
            ElseIf Opt.Item(1).Value = True Then
                Text1.Text = "Dos " & Text1.Text & Time & vbCrLf
                dos
            ElseIf Opt.Item(2).Value = True Then
                Text1.Text = "Tres " & Text1.Text & Time & vbCrLf
                tres
            ElseIf Opt.Item(3).Value = True Then
                Text1.Text = "Cuatro " & Text1.Text & Time & vbCrLf
                cuatro
            ElseIf Opt.Item(4).Value = True Then
                Text1.Text = "Cinco " & Text1.Text & Time & vbCrLf
                cinco
            ElseIf Opt.Item(5).Value = True Then
                Text1.Text = "Seis " & Text1.Text & Time & vbCrLf
                seis
            ElseIf Opt.Item(6).Value = True Then
                Text1.Text = "Siete " & Text1.Text & Time & vbCrLf
                Tres2
            ElseIf Opt.Item(7).Value = True Then
                Text1.Text = "Ocho " & Text1.Text & Time & vbCrLf
                ocho
            ElseIf Opt.Item(8).Value = True Then
                Text1.Text = "Nueve " & Text1.Text & Time & vbCrLf
                Nueve
            ElseIf Opt.Item(9).Value = True Then
                Text1.Text = "Diez " & Text1.Text & Time & vbCrLf
                diez
            End If
            'PARA CUANDO TERMINA EL PROCESO
            If Opt.Item(0).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Uno " & Time
            ElseIf Opt.Item(1).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Dos " & Time
            ElseIf Opt.Item(2).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Tres " & Time
            ElseIf Opt.Item(3).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Cuatro " & Time
            ElseIf Opt.Item(4).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Cinco " & Time
            ElseIf Opt.Item(5).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Seis " & Time
            ElseIf Opt.Item(6).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Siete " & Time
            ElseIf Opt.Item(7).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Ocho " & Time
            ElseIf Opt.Item(8).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Nueve " & Time
            ElseIf Opt.Item(9).Value = True Then
                Text1.Text = vbCrLf & Text1.Text & "Diez " & Time
            End If
                
        MousePointer = 0
                MsgBox "Proceso Terminado Con Exito", vbOKOnly + vbInformation, "Informacion"
End Sub

Public Sub uno()
On Error Resume Next

            'Set RUsuarios = New ADODB.Recordset
            'Call Abrir_Recordset(RUsuarios, "sELECT * fROM detalleentradasinventario D, fichatecnica F where d.bodega = '011' And D.FichaTecnica = F.Esp_Tec And F.TipoInventario = 'MATERIA PRIMA'")

            'If RUsuarios.RecordCount > 0 Then
            '    Do Until RUsuarios.EOF
            '            RUsuarios!Bodega = "024"
            '            RUsuarios.Update
            '
            '            If Err <> 0 Then
            '                MsgBox Err.Description
            '            End If
            '        RUsuarios.MoveNext
            '    Loop
            'End If


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RUsuarios = New ADODB.Recordset
            Call Abrir_Recordset(RUsuarios, "Select * From usuarios")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set Rusuarios2 = New ADODB.Recordset
            Call Abrir_Recordset2(Rusuarios2, "Select * From Usuarios")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RUsuarios.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    Rusuarios2.AddNew
                        Rusuarios2(0) = UCase(RUsuarios(0))
                        Rusuarios2(1) = RUsuarios(1)
                        Rusuarios2(2) = RUsuarios(2)
                        Rusuarios2(3) = RUsuarios(3)
                        Rusuarios2(4) = RUsuarios(4)
                        Rusuarios2(5) = RUsuarios(5)
                        Rusuarios2(6) = RUsuarios(6)
                        Rusuarios2(7) = RUsuarios(7)
                        Rusuarios2(8) = RUsuarios(8)
                        Rusuarios2(9) = RUsuarios(9)
                        Rusuarios2(10) = RUsuarios(10)
                        Rusuarios2(11) = RUsuarios(11)
                        Rusuarios2(12) = RUsuarios(12)
                        Rusuarios2(13) = RUsuarios(13)
                        Rusuarios2(14) = RUsuarios(14)
                        Rusuarios2(15) = RUsuarios(15)
                        Rusuarios2(16) = RUsuarios(16)
                        Rusuarios2(17) = RUsuarios(17)
                        Rusuarios2(18) = RUsuarios(18)
                        Rusuarios2(19) = RUsuarios(19)
                        Rusuarios2(20) = RUsuarios(20)
                        Rusuarios2(21) = RUsuarios(21)
                        Rusuarios2(22) = RUsuarios(22)
                        Rusuarios2(23) = RUsuarios(23)
                        Rusuarios2(24) = RUsuarios(24)
                        Rusuarios2(25) = RUsuarios(25)
                        Rusuarios2(26) = RUsuarios(26)
                        Rusuarios2(27) = RUsuarios(27)
                        Rusuarios2(28) = RUsuarios(28)
                        Rusuarios2(29) = RUsuarios(29)
                        Rusuarios2(30) = RUsuarios(30)
                        Rusuarios2(31) = RUsuarios(31)
                        Rusuarios2(32) = RUsuarios(32)
                        Rusuarios2(33) = RUsuarios(33)
                        Rusuarios2(34) = RUsuarios(34)
                        Rusuarios2(35) = RUsuarios(35)
                        Rusuarios2(36) = RUsuarios(36)
                        Rusuarios2(37) = RUsuarios(37)
                        Rusuarios2(38) = RUsuarios(38)
                        Rusuarios2(39) = RUsuarios(39)
                        Rusuarios2(40) = RUsuarios(40)
                        Rusuarios2(41) = RUsuarios(41)
                        Rusuarios2(42) = RUsuarios(42)
                        Rusuarios2(43) = RUsuarios(43)
                        Rusuarios2(44) = RUsuarios(44)
                        Rusuarios2(45) = RUsuarios(45)
                        Rusuarios2(46) = RUsuarios(46)
                        Rusuarios2(47) = RUsuarios(47)
                        Rusuarios2(48) = RUsuarios(48)
                        Rusuarios2!FechaAlta = RUsuarios!FechaAlta
                        Rusuarios2!FechaUltimoAcceso = RUsuarios!FechaUltimoAcceso
                        Rusuarios2!ContadorAccesos = RUsuarios!ContadorAccesos
                        
                    Rusuarios2.Update
                    If Err <> 0 Then
                        'MsgBox Err.Description & "Usuarios"
                        Err.Clear
                    End If
                RUsuarios.MoveNext
            Loop


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            
            
            'Set RAtributos = New ADODB.Recordset
            'Call Abrir_Recordset(RAtributos, "Select * From Atributos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RAtributos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RAtributos2, "Select * From Atributos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RAtributos.EOF
            '        'AGREGA UN REGISTRO EN ORACLE
            '        RAtributos2.AddNew
            '            RAtributos2(0) = RAtributos(0)
            '            RAtributos2(1) = RAtributos(1)
            '            RAtributos2(2) = RAtributos(2)
            '            RAtributos2(3) = RAtributos(3)
            '            RAtributos2(4) = RAtributos(4)
            '        RAtributos2.Update
           '
           '         If Err <> 0 Then
            '            MsgBox Err.Description & "Atributos"
             '           Err.Clear
             '       End If
           '
            '    RAtributos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RBodegasMateriaPrimaGrupos = New ADODB.Recordset
            'Call Abrir_Recordset(RBodegasMateriaPrimaGrupos, "Select * From BodegasMateriaPrimaGrupos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RBodegasMateriaPrimaGrupos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RBodegasMateriaPrimaGrupos2, "Select * From BodegasMateriaPrimaGrupos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RBodegasMateriaPrimaGrupos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RBodegasMateriaPrimaGrupos2.AddNew
            '            RBodegasMateriaPrimaGrupos2(0) = RBodegasMateriaPrimaGrupos(0)
            '            RBodegasMateriaPrimaGrupos2(1) = RBodegasMateriaPrimaGrupos(1)
            '            RBodegasMateriaPrimaGrupos2(2) = RBodegasMateriaPrimaGrupos(2)
            '        RBodegasMateriaPrimaGrupos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Bodegas MP Grupos"
            '            Err.Clear
            '        End If
            '    RBodegasMateriaPrimaGrupos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RBodegasMateriaPrima = New ADODB.Recordset
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * From BodegasMateriaPrima")
            ''ABRIMOS EL RECORDSET DE ORACLE
            Set RBodegasMateriaPrima2 = New ADODB.Recordset
            Call Abrir_Recordset2(RBodegasMateriaPrima2, "Select * From BodegasInventario")
            ''HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RBodegasMateriaPrima.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RBodegasMateriaPrima2.AddNew
                        RBodegasMateriaPrima2(0) = RBodegasMateriaPrima(0)
                        RBodegasMateriaPrima2(1) = RBodegasMateriaPrima(1)
                        RBodegasMateriaPrima2(2) = RBodegasMateriaPrima(2)
                        RBodegasMateriaPrima2(3) = RBodegasMateriaPrima(3)
                        RBodegasMateriaPrima2(4) = RBodegasMateriaPrima(4)
                        RBodegasMateriaPrima2(5) = RBodegasMateriaPrima(5)
                        RBodegasMateriaPrima2(6) = RBodegasMateriaPrima(6)
                        RBodegasMateriaPrima2(7) = RBodegasMateriaPrima(7)
                        RBodegasMateriaPrima2(8) = RBodegasMateriaPrima(8)
                    RBodegasMateriaPrima2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Bodegas Mp"
                        Err.Clear
                    End If
                RBodegasMateriaPrima.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RBodegasProductoTerminadoGrupos = New ADODB.Recordset
            'Call Abrir_Recordset(RBodegasProductoTerminadoGrupos, "Select * From BodegasProductoTerminadoGrupos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RBodegasProductoTerminadoGrupos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RBodegasProductoTerminadoGrupos2, "Select * From BodegasProductoTerminadoGrupos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RBodegasProductoTerminadoGrupos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RBodegasProductoTerminadoGrupos2.AddNew
            '            RBodegasProductoTerminadoGrupos2(0) = RBodegasProductoTerminadoGrupos(0)
            '            RBodegasProductoTerminadoGrupos2(1) = RBodegasProductoTerminadoGrupos(1)
            '            RBodegasProductoTerminadoGrupos2(2) = RBodegasProductoTerminadoGrupos(2)
            '        RBodegasProductoTerminadoGrupos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Bodegas PT Grupos"
            '            Err.Clear
            '        End If
            '    RBodegasProductoTerminadoGrupos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RBodegasProductoTerminado = New ADODB.Recordset
            Call Abrir_Recordset(RBodegasProductoTerminado, "Select * From BodegasProductoTerminado")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RBodegasProductoTerminado2 = New ADODB.Recordset
            Call Abrir_Recordset2(RBodegasProductoTerminado2, "Select * From BodegasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RBodegasProductoTerminado.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RBodegasProductoTerminado2.AddNew
                        RBodegasProductoTerminado2(0) = RBodegasProductoTerminado(0)
                        RBodegasProductoTerminado2(1) = RBodegasProductoTerminado(1)
                        RBodegasProductoTerminado2(2) = RBodegasProductoTerminado(2)
                        RBodegasProductoTerminado2(3) = RBodegasProductoTerminado(3)
                        RBodegasProductoTerminado2(4) = RBodegasProductoTerminado(4)
                        RBodegasProductoTerminado2(5) = RBodegasProductoTerminado(5)
                        RBodegasProductoTerminado2(6) = RBodegasProductoTerminado(6)
                        RBodegasProductoTerminado2(7) = RBodegasProductoTerminado(7)
                        RBodegasProductoTerminado2(8) = RBodegasProductoTerminado(8)
                    RBodegasProductoTerminado2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Bodegas PT"
                        Err.Clear
                    End If
                RBodegasProductoTerminado.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RClientes = New ADODB.Recordset
            Call Abrir_Recordset(RClientes, "Select * From Clientes")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RClientes2 = New ADODB.Recordset
            Call Abrir_Recordset2(RClientes2, "Select * From Clientes")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RClientes.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RClientes2.AddNew
                        RClientes2(0) = RClientes(0)
                        RClientes2(1) = RClientes(1)
                        RClientes2(2) = RClientes(2)
                        RClientes2(3) = RClientes(3)
                        RClientes2(4) = RClientes(4)
                        RClientes2(5) = RClientes(5)
                        RClientes2(6) = RClientes(6)
                        RClientes2(7) = RClientes(7)
                    RClientes2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Clientes"
                        Err.Clear
                    End If
                RClientes.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RTiposMP = New ADODB.Recordset
            Call Abrir_Recordset(RTiposMP, "Select * From TiposDeMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RTiposMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RTiposMP2, "Select * From FichaTecnicaTipos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RTiposMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RTiposMP2.AddNew
                        RTiposMP2(0) = RTiposMP(0)
                        RTiposMP2(1) = RTiposMP(1)
                        RTiposMP2(2) = RTiposMP(2)
                    RTiposMP2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Tipos Materia Prima"
                        Err.Clear
                    End If
                RTiposMP.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        'ABRIMOS EL RECORDSET DE ACCESS
            Set RTiposMP = New ADODB.Recordset
            Call Abrir_Recordset(RTiposMP, "Select * From FichaTecnicaTipos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RTiposMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RTiposMP2, "Select * From FichaTecnicaTipos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RTiposMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RTiposMP2.AddNew
                        RTiposMP2(0) = RTiposMP(0)
                        RTiposMP2(1) = RTiposMP(1)
                        RTiposMP2(2) = "Erick"
                    RTiposMP2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Ficha Tecnica Tipos"
                        Err.Clear
                    End If
                RTiposMP.MoveNext
            Loop
        
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDefectos = New ADODB.Recordset
            Call Abrir_Recordset(RDefectos, "Select * From Defectos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDefectos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDefectos2, "Select * From Defectos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDefectos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDefectos2.AddNew
                        RDefectos2(0) = RDefectos(0)
                        RDefectos2(1) = RDefectos(1)
                        RDefectos2(2) = RDefectos(2)
                        RDefectos2(3) = "Erick"
                    RDefectos2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Defectos"
                        Err.Clear
                    End If
                RDefectos.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RDocumentos = New ADODB.Recordset
            'Call Abrir_Recordset(RDocumentos, "Select * From Documentos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RDocumentos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RDocumentos2, "Select * From Documentos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RDocumentos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RDocumentos2.AddNew
            '            RDocumentos2(0) = RDocumentos(0)
            '            RDocumentos2(1) = RDocumentos(1)
            '        RDocumentos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Documentos"
            '            Err.Clear
            '        End If
            '    RDocumentos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________

End Sub

Public Sub dos()
On Error Resume Next

'ABRIMOS EL RECORDSET DE ACCESS
            'Set RCobrosProveedor = New ADODB.Recordset
            'Call Abrir_Recordset(RCobrosProveedor, "Select * From CobrosProveedor")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RCobrosProveedor2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RCobrosProveedor2, "Select * From CobrosProveedor")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RCobrosProveedor.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RCobrosProveedor2.AddNew
            '            RCobrosProveedor2!fecha = UCase(RCobrosProveedor!fecha)
            '            RCobrosProveedor2!Proveedor = RCobrosProveedor!Proveedor
            '            RCobrosProveedor2!FichaTecnica = UCase(RCobrosProveedor!FichaTecnica)
            '            RCobrosProveedor2!Defecto = RCobrosProveedor!Defecto
            '            RCobrosProveedor2!Boleta = UCase(RCobrosProveedor!Boleta)
            '            RCobrosProveedor2!FechaRevision = RCobrosProveedor!FechaRevision
            '            RCobrosProveedor2!URevisadas = RCobrosProveedor!URevisadas
            '            RCobrosProveedor2!UNoConformes = RCobrosProveedor!UNoConformes
             ''           RCobrosProveedor2!CostoxUnidad = RCobrosProveedor!CostoxUnidad
            '            RCobrosProveedor2!HorasHombre = RCobrosProveedor!HorasHombre
            '            RCobrosProveedor2!CostoxHora = RCobrosProveedor!CostoxHora
            '            RCobrosProveedor2!TazaCambio = RCobrosProveedor!TazaCambio
            '            RCobrosProveedor2!Usuario = RCobrosProveedor!Usuario
            '            RCobrosProveedor2!Serie = RCobrosProveedor!Serie
                        
            '        RCobrosProveedor2.Update
            '        If Err <> 0 Then
                        'MsgBox Err.Description & "cobros proveedores"
            '            Err.Clear
            '        End If
            '    RCobrosProveedor.MoveNext
            'Loop


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REmpleados = New ADODB.Recordset
            Call Abrir_Recordset(REmpleados, "Select * From Empleados")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REmpleados2 = New ADODB.Recordset
            Call Abrir_Recordset2(REmpleados2, "Select * From Empleados")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            
            Do Until REmpleados.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    'Conexion2.Execute "Update Empleados Set Telefono = '" & REmpleados!Telefono & "', FechaNacimiento = To_Date('" & REmpleados!FechaNacimiento & "', 'dd/mm/yyyy') Where Codigo = '" & REmpleados!Codigo & "'"
                    
                    REmpleados2.AddNew
                        REmpleados2(0) = UCase(REmpleados(0))
                        REmpleados2(1) = REmpleados(1)
                        REmpleados2(2) = UCase(REmpleados(2))
                        REmpleados2(3) = REmpleados(3)
                        REmpleados2(4) = UCase(REmpleados(4))
                        REmpleados2(5) = REmpleados(5)
                        'REmpleados2(6) = REmpleados(6)
                        'REmpleados2(7) = REmpleados(7)
                        'REmpleados2(8) = REmpleados(8)
                        'REmpleados2(9) = REmpleados(9)
                        'REmpleados2(10) = REmpleados(10)
                        'REmpleados2(11) = REmpleados(11)
                        'REmpleados2(12) = REmpleados(12)
                        'REmpleados2(12) = REmpleados(13)
                        'REmpleados2(14) = REmpleados(14)
                        'REmpleados2(15) = REmpleados(15)
                        'REmpleados2(16) = REmpleados(16)
                        'REmpleados2(17) = UCase(REmpleados(17))
                        'REmpleados2(18) = REmpleados(18)
                        'REmpleados2(19) = REmpleados(19)
                        'REmpleados2(20) = REmpleados(20)
                        'REmpleados2(21) = REmpleados(21)
                        'REmpleados2(22) = REmpleados(22)
                        'REmpleados2(23) = REmpleados(23)
                    REmpleados2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Empleados"
                        Err.Clear
                    End If
                REmpleados.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosCursos = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosCursos, "Select * From EmpleadosCursos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosCursos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosCursos2, "Select * From EmpleadosCursos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosCursos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
             '       REmpleadosCursos2.AddNew
             '           REmpleadosCursos2(0) = UCase(REmpleadosCursos(0))
             '           REmpleadosCursos2(1) = REmpleadosCursos(1)
             '           REmpleadosCursos2(2) = REmpleadosCursos(2)
             '           REmpleadosCursos2(3) = REmpleadosCursos(3)
             '       REmpleadosCursos2.Update
             '       If Err <> 0 Then
             '           MsgBox Err.Description & "Empleados Cursos"
             '           Err.Clear
             '       End If
             '   REmpleadosCursos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosDepartamentos = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosDepartamentos, "Select * From EmpleadosDepartamentos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosDepartamentos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosDepartamentos2, "Select * From EmpleadosDepartamentos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosDepartamentos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosDepartamentos2.AddNew
            '            REmpleadosDepartamentos2(0) = UCase(REmpleadosDepartamentos(0))
            '            REmpleadosDepartamentos2(1) = REmpleadosDepartamentos(1)
            '            REmpleadosDepartamentos2(2) = REmpleadosDepartamentos(2)
            '        REmpleadosDepartamentos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Departamentos"
            '           Err.Clear
            '        End If
            '    REmpleadosDepartamentos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosEscolaridad = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosEscolaridad, "Select * From EmpleadosEscolaridad")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosEscolaridad2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosEscolaridad2, "Select * From EmpleadosEscolaridad")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosEscolaridad.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosEscolaridad2.AddNew
            '            REmpleadosEscolaridad2(0) = UCase(REmpleadosEscolaridad(0))
            '            REmpleadosEscolaridad2(1) = REmpleadosEscolaridad(1)
            '            REmpleadosEscolaridad2(2) = REmpleadosEscolaridad(2)
            '        REmpleadosEscolaridad2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Escolaridad"
            '            Err.Clear
            '        End If
            '    REmpleadosEscolaridad.MoveNext
            'Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosFaltas = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosFaltas, "Select * From EmpleadosFaltas")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosFaltas2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosFaltas2, "Select * From EmpleadosFaltas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosFaltas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosFaltas2.AddNew
            '            REmpleadosFaltas2(0) = UCase(REmpleadosFaltas(0))
            '            REmpleadosFaltas2(1) = REmpleadosFaltas(1)
            '            REmpleadosFaltas2(2) = REmpleadosFaltas(2)
            '        REmpleadosFaltas2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Faltas"
            '            Err.Clear
            '        End If
            '    REmpleadosFaltas.MoveNext
            'Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosGrupos = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosGrupos, "Select * From EmpleadosGrupos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosGrupos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosGrupos2, "Select * From EmpleadosGrupos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosGrupos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosGrupos2.AddNew
            '            REmpleadosGrupos2(0) = UCase(REmpleadosGrupos(0))
            '            REmpleadosGrupos2(1) = REmpleadosGrupos(1)
            '            REmpleadosGrupos2(2) = REmpleadosGrupos(2)
            '        REmpleadosGrupos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Grupos"
            '            Err.Clear
            ''        End If
            '    REmpleadosGrupos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosHabilidades = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosHabilidades, "Select * From EmpleadosHabilidades")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosHabilidades2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosHabilidades2, "Select * From EmpleadosHabilidades")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosHabilidades.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosHabilidades2.AddNew
            '            REmpleadosHabilidades2(0) = UCase(REmpleadosHabilidades(0))
            '            REmpleadosHabilidades2(1) = REmpleadosHabilidades(1)
            '            REmpleadosHabilidades2(2) = REmpleadosHabilidades(2)
            '            REmpleadosHabilidades2(3) = REmpleadosHabilidades(3)
            '        REmpleadosHabilidades2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Habilidades"
            '            Err.Clear
            '        End If
            '    REmpleadosHabilidades.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
           
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosPuestos = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosPuestos, "Select * From EmpleadosPuestos")
            'ABRIMOS EL RECORDSET DE ORACLE
           ' Set REmpleadosPuestos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosPuestos2, "Select * From EmpleadosPuestos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosPuestos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosPuestos2.AddNew
            '            REmpleadosPuestos2(0) = UCase(REmpleadosPuestos(0))
            '            REmpleadosPuestos2(1) = REmpleadosPuestos(1)
            '            REmpleadosPuestos2(2) = REmpleadosPuestos(2)
            '            REmpleadosPuestos2(3) = REmpleadosPuestos(3)
            '        REmpleadosPuestos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Puestos"
            '            Err.Clear
            '        End If
            '    REmpleadosPuestos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosHabEmp = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosHabEmp, "Select * From EmpleadosHabilidadesEmpleado")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosHabEmp2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosHabEmp2, "Select * From EmpleadosHabilidadesEmpleado")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosHabEmp.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosHabEmp2.AddNew
            '           REmpleadosHabEmp2(0) = UCase(REmpleadosHabEmp(0))
            '            REmpleadosHabEmp2(1) = UCase(REmpleadosHabEmp(1))
            '            REmpleadosHabEmp2(2) = REmpleadosHabEmp(2)
            '       REmpleadosHabEmp2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Habilidades Empleados"
            '            Err.Clear
            '        End If
            '    REmpleadosHabEmp.MoveNext
            'Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REmpleadosHabilidadesPuestos = New ADODB.Recordset
            'Call Abrir_Recordset(REmpleadosHabilidadesPuestos, "Select * From EmpleadosHabilidadesPuesto")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REmpleadosHabilidadesPuestos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REmpleadosHabilidadesPuestos2, "Select * From EmpleadosHabilidadesPuesto")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REmpleadosHabilidadesPuestos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REmpleadosHabilidadesPuestos2.AddNew
            '            REmpleadosHabilidadesPuestos2(0) = UCase(REmpleadosHabilidadesPuestos(0))
            '            REmpleadosHabilidadesPuestos2(1) = UCase(REmpleadosHabilidadesPuestos(1))
            '            REmpleadosHabilidadesPuestos2(2) = REmpleadosHabilidadesPuestos(2)
            '        REmpleadosHabilidadesPuestos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Empleados Habilidades Puestos"
            '            Err.Clear
            '        End If
            '    REmpleadosHabilidadesPuestos.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RVariablesDescripcion = New ADODB.Recordset
            Call Abrir_Recordset(RVariablesDescripcion, "Select * From VariablesDescripcion")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RVariablesDescripcion2 = New ADODB.Recordset
            Call Abrir_Recordset2(RVariablesDescripcion2, "Select * From VariablesDescripcion")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RVariablesDescripcion.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RVariablesDescripcion2.AddNew
                        RVariablesDescripcion2(0) = RVariablesDescripcion(0)
                        RVariablesDescripcion2(1) = RVariablesDescripcion(1)
                        RVariablesDescripcion2(2) = "GEOVANNI"
                    RVariablesDescripcion2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Variables Descripcion"
                        Err.Clear
                    End If
                RVariablesDescripcion.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RFichaTecnica = New ADODB.Recordset
            Call Abrir_Recordset(RFichaTecnica, "Select * From FichaTecnica")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RFichaTecnica2 = New ADODB.Recordset
            Call Abrir_Recordset2(RFichaTecnica2, "Select * From FichaTecnica")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RFichaTecnica.EOF
            
                    
                    'AGREGA UN REGISTRO EN ORACLE
                    RFichaTecnica2.AddNew
                        RFichaTecnica2(0) = UCase(RFichaTecnica(0))
                        RFichaTecnica2(1) = RFichaTecnica(1)
                        RFichaTecnica2(2) = UCase(RFichaTecnica(2))
                        RFichaTecnica2(3) = RFichaTecnica(3)
                        RFichaTecnica2(4) = RFichaTecnica(4)
                        RFichaTecnica2(5) = RFichaTecnica(5)
                        RFichaTecnica2(6) = RFichaTecnica(6)
                        If IsNull(RFichaTecnica(7)) Then
                            RFichaTecnica2(7) = "VARIOS"
                        Else
                            RFichaTecnica2(7) = RFichaTecnica(7)
                        End If
                        If IsNull(RFichaTecnica(8)) Then
                            RFichaTecnica2(8) = "VARIOS"
                        Else
                            RFichaTecnica2(8) = UCase(RFichaTecnica(8))
                        End If
                        RFichaTecnica2(9) = RFichaTecnica(9)
                        RFichaTecnica2(10) = RFichaTecnica(10)
                        RFichaTecnica2(11) = RFichaTecnica(11)
                        RFichaTecnica2(12) = RFichaTecnica(12)
                        RFichaTecnica2(12) = RFichaTecnica(13)
                        RFichaTecnica2(14) = RFichaTecnica(14)
                        RFichaTecnica2(15) = RFichaTecnica(15)
                        RFichaTecnica2(16) = RFichaTecnica(16)
                        RFichaTecnica2(17) = "PRODUCTO TERMINADO"
                        RFichaTecnica2(18) = "0"
                    RFichaTecnica2.Update
                    If Err = -2147467259 Then
                        MsgBox Err.Description & "Ficha Tecnica"
                        Err.Clear
                    ElseIf Err = -2147217873 Then
 '                       MsgBox Err.Description & "Ficha Tecnica"
                    ElseIf Err <> 0 And Err <> -2147217873 Then
                        MsgBox Err.Description & "Ficha Tecnica"
                        Err.Clear
                    End If
                RFichaTecnica.MoveNext
                'RFichaTecnica.MovePrevious
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            
            Set RCorrelativosMP = New ADODB.Recordset
            Call Abrir_Recordset(RCorrelativosMP, "Select * From CorrelativosMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RCorrelativosMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RCorrelativosMP2, "Select * From FichaTecnica")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RCorrelativosMP.EOF
                    
                    If RCorrelativosMP!TipoDeMateriaPrima = "013" Then
                        'NO HAY QUE ARGREGAR PRODUCTO TERMINADO
                    Else
                            'AGREGA UN REGISTRO EN ORACLE
                            RCorrelativosMP2.AddNew
                                RCorrelativosMP2!Esp_Tec = UCase(RCorrelativosMP!CodigoMateriaPrima)
                                RCorrelativosMP2!Descrip = RCorrelativosMP!Descripcion
                                'If RCorrelativosMP!TipoDeMateriaPrima = "001" Then
                                '    RCorrelativosMP2!Tipo = "30"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "002" Then
                                '    RCorrelativosMP2!Tipo = "31"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "003" Then
                                '    RCorrelativosMP2!Tipo = "32"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "004" Then
                                '    RCorrelativosMP2!Tipo = "33"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "005" Then
                                '    RCorrelativosMP2!Tipo = "34"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "006" Then
                                '    RCorrelativosMP2!Tipo = "35"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "007" Then
                                '    RCorrelativosMP2!Tipo = "36"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "008" Then
                                '    RCorrelativosMP2!Tipo = "37"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "009" Then
                                '    RCorrelativosMP2!Tipo = "38"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "010" Then
                                '    RCorrelativosMP2!Tipo = "39"
                                'ElseIf RCorrelativosMP!TipoDeMateriaPrima = "013" Then
                                'End If
                                RCorrelativosMP2!Tipo = RCorrelativosMP!TipoDeMateriaPrima
                                'RCorrelativosMP2!Diametro = 0
                                'RCorrelativosMP2!Capacida = 0
                                'RCorrelativosMP2!Altura = 0
                                RCorrelativosMP2!Envases = 0
                                RCorrelativosMP2!Nombre_Comercial = ""
                                RCorrelativosMP2!Imp_Defe = 0
                                RCorrelativosMP2!Imp_Cali = 0
                                RCorrelativosMP2!Atributos = "VARIOS"
                                RCorrelativosMP2!Variables = "VARIOS"
                                'RCorrelativosMP2!Foto1 = ""
                                RCorrelativosMP2!Origen = "EXTERNO"
                                RCorrelativosMP2!Usuario = "Erick"
                                RCorrelativosMP2!unidadMedida = RCorrelativosMP!unidadMedida
                                RCorrelativosMP2!MaterialEmpaque = "BULTO"
                                RCorrelativosMP2!Activa = "-1"
                                RCorrelativosMP2!PesoxUnidad = RCorrelativosMP!PesoxUnidad
                                RCorrelativosMP2!UnidadesxLamina = RCorrelativosMP!CuerposPorLamina
                                RCorrelativosMP2!UnidadesxCaja = 0
                                RCorrelativosMP2!TipoInventario = "MATERIA PRIMA"
                                RCorrelativosMP2!Espesor = RCorrelativosMP!Espesor
                            RCorrelativosMP2.Update
                    End If
                            
                            
                    If Err = -2147467259 Then
                        'MsgBox Err.Description & "Correlativos MP"
                        Err.Clear
                    ElseIf Err = -2147217873 Then
                        MsgBox Err.Description & "Correlativos MP"
                        Err.Clear
                    ElseIf Err <> -2147217873 And Err <> 0 Then
                        MsgBox Err.Description & "Correlativos MP"
                        Err.Clear
                    End If
                RCorrelativosMP.MoveNext
            Loop
'______________________________________________________________________________________________________________________


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RFichaTecnicaConMateriaPrima = New ADODB.Recordset
            Call Abrir_Recordset(RFichaTecnicaConMateriaPrima, "Select * From FichaTecnicaConMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RFichaTecnicaConMateriaPrima2 = New ADODB.Recordset
            Call Abrir_Recordset2(RFichaTecnicaConMateriaPrima2, "Select * From FichaTecnicaConMateriaPrima")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RFichaTecnicaConMateriaPrima.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RFichaTecnicaConMateriaPrima2.AddNew
                        RFichaTecnicaConMateriaPrima2(0) = UCase(RFichaTecnicaConMateriaPrima(0))
                        RFichaTecnicaConMateriaPrima2(1) = UCase(RFichaTecnicaConMateriaPrima(1))
                        RFichaTecnicaConMateriaPrima2(2) = RFichaTecnicaConMateriaPrima(2)
                        RFichaTecnicaConMateriaPrima2(3) = RFichaTecnicaConMateriaPrima(3)
                        RFichaTecnicaConMateriaPrima2(4) = RFichaTecnicaConMateriaPrima(4)
                    RFichaTecnicaConMateriaPrima2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Ficha Tecnica Con Materia Prima"
                        Err.Clear
                    End If
                RFichaTecnicaConMateriaPrima.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RLineas = New ADODB.Recordset
            Call Abrir_Recordset(RLineas, "Select * From Lineas")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RLineas2 = New ADODB.Recordset
            Call Abrir_Recordset2(RLineas2, "Select * From Lineas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RLineas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RLineas2.AddNew
                        RLineas2(0) = RLineas(0)
                        RLineas2(1) = RLineas(1)
                        RLineas2(2) = RLineas(2)
                        RLineas2(3) = RLineas(3)
                        RLineas2(4) = RLineas(4)
                        RLineas2(5) = RLineas(5)
                        RLineas2(6) = RLineas(6)
                        RLineas2(7) = RLineas(7)
                        RLineas2(8) = RLineas(8)
                        RLineas2(9) = RLineas(9)
                        RLineas2(10) = RLineas(10)
                   RLineas2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Lineas"
                        Err.Clear
                    End If
                RLineas.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RLineasBultos = New ADODB.Recordset
            Call Abrir_Recordset(RLineasBultos, "Select * From LineasBultos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RLineasBultos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RLineasBultos2, "Select * From LineasBultos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RLineasBultos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RLineasBultos2.AddNew
                        RLineasBultos2(0) = RLineasBultos(0)
                        RLineasBultos2(1) = RLineasBultos(1)
                        RLineasBultos2(2) = RLineasBultos(2)
                        RLineasBultos2(3) = RLineasBultos(3)
                        RLineasBultos2(4) = RLineasBultos(4)
                    RLineasBultos2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Lineas Bultos"
                        Err.Clear
                    End If
                RLineasBultos.MoveNext
            Loop
'______________________________________________________________________________________________________________________

End Sub

Public Sub tres()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RInventario = New ADODB.Recordset
            Call Abrir_Recordset(RInventario, "Select * From Inventario")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RInventario2 = New ADODB.Recordset
            Call Abrir_Recordset2(RInventario2, "Select * From Inventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RInventario.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    If RInventario(1) = "208fo-04-00" Or RInventario(1) = "211300-03" Or RInventario(1) = "211300-04" Or RInventario(1) = "211300-10" Then
                    Else
                        RInventario2.AddNew
                            RInventario2(0) = RInventario(0)
                            RInventario2(1) = UCase(RInventario(1))
                            RInventario2(2) = UCase(RInventario(2))
                            RInventario2(3) = RInventario(3)
                            RInventario2(4) = RInventario(4)
                        RInventario2.Update
                    End If
                    If Err <> 0 Then
                        'MsgBox Err.Description & "Inventario"
                        Err.Clear
                    End If
                RInventario.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RTurnos = New ADODB.Recordset
            'Call Abrir_Recordset(RTurnos, "Select * From Turnos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RTurnos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RTurnos2, "Select * From Turnos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RTurnos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RTurnos2.AddNew
            '            RTurnos2(0) = RTurnos(0)
            '            RTurnos2(1) = RTurnos(1)
            '            RTurnos2(2) = RTurnos(2)
            '        RTurnos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Turnos"
            '            Err.Clear
            '        End If
            '    RTurnos.MoveNext
            'Loop

            

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RParosGrupos = New ADODB.Recordset
            Call Abrir_Recordset(RParosGrupos, "Select * From ParosGrupos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RParosGrupos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RParosGrupos2, "Select * From ParosGrupos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RParosGrupos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RParosGrupos2.AddNew
                        RParosGrupos2(0) = UCase(RParosGrupos(0))
                        RParosGrupos2(1) = RParosGrupos(1)
                        RParosGrupos2(2) = RParosGrupos(2)
                    RParosGrupos2.Update
                    If Err = -2147217873 Then
                    ElseIf Err <> 0 Then
                        MsgBox Err.Description & "Paros Grupos"
                        Err.Clear
                    End If
                RParosGrupos.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RParos = New ADODB.Recordset
            Call Abrir_Recordset(RParos, "Select * From Paros")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RParos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RParos2, "Select * From Paros")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RParos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RParos2.AddNew
                        RParos2(0) = UCase(RParos(0))
                        RParos2(1) = RParos(1)
                        RParos2(2) = UCase(RParos(2))
                        RParos2(3) = RParos(4)
                        RParos2(4) = UCase(RParos(5))
                    RParos2.Update
                    If Err.Number = -2147217873 Then
                    
                    ElseIf Err <> 0 Then
                        MsgBox Err.Description & "Paros"
                        Err.Clear
                    End If
                RParos.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RPasadas = New ADODB.Recordset
            'Call Abrir_Recordset(RPasadas, "Select * From Pasadas")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RPasadas2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RPasadas2, "Select * From Pasadas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
           ' Do Until RPasadas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RPasadas2.AddNew
            '            RPasadas2(0) = UCase(RPasadas(0))
            '            RPasadas2(1) = RPasadas(1)
            '            RPasadas2(2) = RPasadas(2)
            '        RPasadas2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Pasadas"
            '            Err.Clear
            '        End If
            '    RPasadas.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RPedidosProveedoresPorcentajeNoConforme = New ADODB.Recordset
            Call Abrir_Recordset(RPedidosProveedoresPorcentajeNoConforme, "Select * From PedidosProveedoresPorcentajeNoConforme")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RPedidosProveedoresPorcentajeNoConforme2 = New ADODB.Recordset
            Call Abrir_Recordset2(RPedidosProveedoresPorcentajeNoConforme2, "Select * From PedidosProveedoresPorcentajeNo")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RPedidosProveedoresPorcentajeNoConforme.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RPedidosProveedoresPorcentajeNoConforme2.AddNew
                        RPedidosProveedoresPorcentajeNoConforme2(0) = RPedidosProveedoresPorcentajeNoConforme(0)
                        RPedidosProveedoresPorcentajeNoConforme2(1) = RPedidosProveedoresPorcentajeNoConforme(1)
                        RPedidosProveedoresPorcentajeNoConforme2(2) = RPedidosProveedoresPorcentajeNoConforme(2)
                        RPedidosProveedoresPorcentajeNoConforme2(3) = UCase(RPedidosProveedoresPorcentajeNoConforme(3))
                        RPedidosProveedoresPorcentajeNoConforme2(4) = RPedidosProveedoresPorcentajeNoConforme(4)
                        RPedidosProveedoresPorcentajeNoConforme2(5) = RPedidosProveedoresPorcentajeNoConforme(5)
                    RPedidosProveedoresPorcentajeNoConforme2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Pedidos Proveedores % No Conforme"
                        Err.Clear
                    End If
                RPedidosProveedoresPorcentajeNoConforme.MoveNext
            Loop
'______________________________________________________________________________________________________________________
        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RProcesosMP = New ADODB.Recordset
            'Call Abrir_Recordset(RProcesosMP, "Select * From ProcesosMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RProcesosMP2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RProcesosMP2, "Select * From ProcesosMateriaPrima")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RProcesosMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RProcesosMP2.AddNew
            '            RProcesosMP2(0) = UCase(RProcesosMP(0))
            '            RProcesosMP2(1) = RProcesosMP(1)
            '            RProcesosMP2(2) = RProcesosMP(2)
            '            RProcesosMP2(3) = RProcesosMP(3)
            '            RProcesosMP2(4) = RProcesosMP(4)
            '        RProcesosMP2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Procesos Materia Prima"
            '            Err.Clear
            '        End If
            '    RProcesosMP.MoveNext
            'Loop
'______________________________________________________________________________________________________________________

End Sub

Public Sub cuatro()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RProduccion = New ADODB.Recordset
            'Call Abrir_Recordset(RProduccion, "Select * From Produccion where month(Fec_Prd) = 06 and Year(fec_prd) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RProduccion2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RProduccion2, "Select * From Produccion")
            ''HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RProduccion.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RProduccion2.AddNew
            '            RProduccion2(0) = RProduccion(0)
            '            RProduccion2(1) = UCase(RProduccion(1))
            '            RProduccion2(2) = UCase(RProduccion(2))
            '            RProduccion2(3) = UCase(RProduccion(3))
            '            RProduccion2(4) = RProduccion(4)
            '            RProduccion2(5) = RProduccion(5)
            '            RProduccion2(6) = RProduccion(6)
            '            RProduccion2(7) = UCase(RProduccion(7))
            '            RProduccion2(8) = UCase(RProduccion(8))
            '            RProduccion2(9) = RProduccion(9)
            '            RProduccion2(10) = UCase(RProduccion(10))
            '            RProduccion2(11) = UCase(RProduccion(11))
            '            RProduccion2(12) = UCase(RProduccion(12))
            '            RProduccion2(12) = UCase(RProduccion(13))
            '            RProduccion2(14) = UCase(RProduccion(14))
            '            RProduccion2(15) = UCase(RProduccion(15))
            '        RProduccion2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Produccion"
            '            Err.Clear
            '        End If
            '    RProduccion.MoveNext
            'Loop
            
           ' MsgBox "YA"
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RProduccionConDefectos = New ADODB.Recordset
            'Call Abrir_Recordset(RProduccionConDefectos, "Select * From ProduccionConDefectos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RProduccionConDefectos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RProduccionConDefectos2, "Select * From ProduccionConDefectos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RProduccionConDefectos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RProduccionConDefectos2.AddNew
            '            RProduccionConDefectos2(0) = RProduccionConDefectos(0)
            '            RProduccionConDefectos2(1) = UCase(RProduccionConDefectos(1))
            '            RProduccionConDefectos2(2) = UCase(RProduccionConDefectos(2))
            '            RProduccionConDefectos2(3) = RProduccionConDefectos(3)
            '            RProduccionConDefectos2(4) = UCase(RProduccionConDefectos(4))
            '            RProduccionConDefectos2(5) = RProduccionConDefectos(5)
            '        RProduccionConDefectos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Produccion Con Defectos"
            '            Err.Clear
            '        End If
            '    RProduccionConDefectos.MoveNext
            'Loop
            
            'MsgBox "ya"
            
            
'______________________________________________________________________________________________________________________
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RProduccionConMP = New ADODB.Recordset
           Call Abrir_Recordset(RProduccionConMP, "Select * From ProduccionConMateriaPrima Where Month(Fec_Prd) = 10 And Year(Fec_Prd) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RProduccionConMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RProduccionConMP2, "Select * From ProduccionConMateriaPrima")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RProduccionConMP.EOF
            
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima E, DetalleEntradasMateriaPrima D Where E.Documento = D.Documento And D.Codigo = '" & RProduccionConMP!CodigoMateriaPrima & "' And D.NumeroIngreso = " & RProduccionConMP!Bulto)
                            
                            
                            If RBuscaFechaEntrada.RecordCount > 0 Then
                                
                            
                        
                                    'AGREGA UN REGISTRO EN ORACLE
                                    RProduccionConMP2.AddNew
                                        RProduccionConMP2!fec_prd = RProduccionConMP!fec_prd
                                        RProduccionConMP2!Linea = UCase(RProduccionConMP!Linea)
                                        RProduccionConMP2!Esp_Tec = UCase(RProduccionConMP!Esp_Tec)
                                        RProduccionConMP2!Tarima = RProduccionConMP!Tarima
                                        RProduccionConMP2!CodigoMateriaPrima = UCase(RProduccionConMP!CodigoMateriaPrima)
                                        RProduccionConMP2!Bulto = RProduccionConMP!Bulto
                                        RProduccionConMP2!Fechaproduccion = RBuscaFechaEntrada(0)
                                        RProduccionConMP2!LineaProduccion = "77"
                                    RProduccionConMP2.Update
                                    If Err <> 0 Then
                                        MsgBox Err.Description & "Produccion Con Materia Prima"
                                        Err.Clear
                                    End If
                            Else
                                'MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            End If
                RProduccionConMP.MoveNext
            Loop
           
      
      
      
      
      
      
      
      
      
      
      MsgBox "ya"
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RProduccionLiberada = New ADODB.Recordset
            'Call Abrir_Recordset(RProduccionLiberada, "Select * From ProduccionLiberada Where Year(Fec_Prd) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RProduccionLiberada2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RProduccionLiberada2, "Select * From ProduccionLiberada")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RProduccionLiberada.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RProduccionLiberada2.AddNew
            '            RProduccionLiberada2(0) = RProduccionLiberada(0)
            '            RProduccionLiberada2(1) = UCase(RProduccionLiberada(1))
            '            RProduccionLiberada2(2) = UCase(RProduccionLiberada(2))
            '            RProduccionLiberada2(3) = UCase(RProduccionLiberada(3))
            '            RProduccionLiberada2(4) = RProduccionLiberada(4)
            '            RProduccionLiberada2(5) = RProduccionLiberada(5)
            '            RProduccionLiberada2(6) = RProduccionLiberada(6)
            '            RProduccionLiberada2(7) = UCase(RProduccionLiberada(7))
            '            RProduccionLiberada2(8) = UCase(RProduccionLiberada(8))
            '            RProduccionLiberada2(9) = UCase(RProduccionLiberada(9))
            '            RProduccionLiberada2(10) = UCase(RProduccionLiberada(10))
            '            RProduccionLiberada2(11) = UCase(RProduccionLiberada(11))
            '            RProduccionLiberada2(12) = UCase(RProduccionLiberada(12))
            '            RProduccionLiberada2(12) = UCase(RProduccionLiberada(13))
            '            RProduccionLiberada2(14) = UCase(RProduccionLiberada(14))
            '        RProduccionLiberada2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Produccion Liberada"
            '            Err.Clear
            '        End If
            '    RProduccionLiberada.MoveNext
            '
            'Loop
            
               
'MsgBox "ya"
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
 '           Set RProduccionLiberadaConDefectos = New ADODB.Recordset
 '           Call Abrir_Recordset(RProduccionLiberadaConDefectos, "Select * From ProduccionLiberadaConDefectos")
 '           'ABRIMOS EL RECORDSET DE ORACLE
 '           Set RProduccionLiberadaConDefectos2 = New ADODB.Recordset
 '           Call Abrir_Recordset2(RProduccionLiberadaConDefectos2, "Select * From ProduccionLiberadaConDefectos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
 '           Do Until RProduccionLiberadaConDefectos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
 '                   RProduccionLiberadaConDefectos2.AddNew
 '                       RProduccionLiberadaConDefectos2(0) = RProduccionLiberadaConDefectos(0)
 '                       RProduccionLiberadaConDefectos2(1) = UCase(RProduccionLiberadaConDefectos(1))
 '                       RProduccionLiberadaConDefectos2(2) = UCase(RProduccionLiberadaConDefectos(2))
 '                       RProduccionLiberadaConDefectos2(3) = RProduccionLiberadaConDefectos(3)
 '                       RProduccionLiberadaConDefectos2(4) = RProduccionLiberadaConDefectos(4)
 '                       RProduccionLiberadaConDefectos2(5) = UCase(RProduccionLiberadaConDefectos(5))
 '                       RProduccionLiberadaConDefectos2(6) = UCase(RProduccionLiberadaConDefectos(6))
 '                       RProduccionLiberadaConDefectos2(7) = RProduccionLiberadaConDefectos(7)
 '                       RProduccionLiberadaConDefectos2(8) = UCase(RProduccionLiberadaConDefectos(8))
  '                      RProduccionLiberadaConDefectos2(9) = RProduccionLiberadaConDefectos(9)
  '                  RProduccionLiberadaConDefectos2.Update
   '                 If Err <> 0 Then
   '                     MsgBox Err.Description & "Produccion Liberada Con Defectos"
    '                    Err.Clear
    '                End If
    '            RProduccionLiberadaConDefectos.MoveNext
    '        Loop
            
    '        MsgBox "ya"
           
           
'______________________________________________________________________________________________________________________
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
     '       Set RProduccionLiberadaConTarimas = New ADODB.Recordset
     '       Call Abrir_Recordset(RProduccionLiberadaConTarimas, "Select * From ProduccionLiberadaConTarimas Where Year(Fec_Prd) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
     '       Set RProduccionLiberadaConTarimas2 = New ADODB.Recordset
     '       Call Abrir_Recordset2(RProduccionLiberadaConTarimas2, "Select * From ProduccionLiberadaConTarimas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
     '       Do Until RProduccionLiberadaConTarimas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
     '               RProduccionLiberadaConTarimas2.AddNew
     '                   RProduccionLiberadaConTarimas2(0) = RProduccionLiberadaConTarimas(0)
     '                   RProduccionLiberadaConTarimas2(1) = UCase(RProduccionLiberadaConTarimas(1))
     '                   RProduccionLiberadaConTarimas2(2) = UCase(RProduccionLiberadaConTarimas(2))
     '                   RProduccionLiberadaConTarimas2(3) = RProduccionLiberadaConTarimas(3)
     '                   RProduccionLiberadaConTarimas2(4) = RProduccionLiberadaConTarimas(4)
     '                   RProduccionLiberadaConTarimas2(5) = UCase(RProduccionLiberadaConTarimas(5))
     '                   RProduccionLiberadaConTarimas2(6) = UCase(RProduccionLiberadaConTarimas(6))
     '                   RProduccionLiberadaConTarimas2(7) = RProduccionLiberadaConTarimas(7)
     '                   RProduccionLiberadaConTarimas2(8) = UCase(RProduccionLiberadaConTarimas(8))
     '                   RProduccionLiberadaConTarimas2(9) = RProduccionLiberadaConTarimas(9)
     '                   RProduccionLiberadaConTarimas2(10) = RProduccionLiberadaConTarimas(10)
     '                   RProduccionLiberadaConTarimas2(11) = RProduccionLiberadaConTarimas(11)
     '                   RProduccionLiberadaConTarimas2(12) = RProduccionLiberadaConTarimas(12)
     '                   RProduccionLiberadaConTarimas2(13) = RProduccionLiberadaConTarimas(13)
     '                   RProduccionLiberadaConTarimas2(14) = UCase(RProduccionLiberadaConTarimas(14))
     '               RProduccionLiberadaConTarimas2.Update
     '               If Err = -2147217873 Then
     '               ElseIf Err <> 0 Then
     '                   MsgBox Err.Description & "Produccion Liberada Con Tarimas"
     '                   Err.Clear
     '               End If
     '           RProduccionLiberadaConTarimas.MoveNext
     '       Loop
            
     '   MsgBox "ya"
'______________________________________________________________________________________________________________________

End Sub

Public Sub cinco()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RProveedoresGrupos = New ADODB.Recordset
            'Call Abrir_Recordset(RProveedoresGrupos, "Select * From ProveedoresGrupos")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RProveedoresGrupos2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RProveedoresGrupos2, "Select * From ProveedoresGrupos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RProveedoresGrupos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RProveedoresGrupos2.AddNew
            '            RProveedoresGrupos2(0) = RProveedoresGrupos(0)
            '            RProveedoresGrupos2(1) = RProveedoresGrupos(1)
            '            RProveedoresGrupos2(2) = UCase(RProveedoresGrupos(2))
            '        RProveedoresGrupos2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Proveedores Grupos"
            '            Err.Clear
            '        End If
            '    RProveedoresGrupos.MoveNext
            'Loop
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RProveedores = New ADODB.Recordset
            Call Abrir_Recordset(RProveedores, "Select * From Proveedores")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RProveedores2 = New ADODB.Recordset
            Call Abrir_Recordset2(RProveedores2, "Select * From Proveedores")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RProveedores.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RProveedores2.AddNew
                        RProveedores2(0) = UCase(RProveedores(0))
                        RProveedores2(1) = UCase(RProveedores(1))
                        RProveedores2(2) = UCase(RProveedores(2))
                        RProveedores2(3) = UCase(RProveedores(3))
                        RProveedores2(4) = UCase(RProveedores(4))
                        RProveedores2(5) = UCase(RProveedores(5))
                        RProveedores2(6) = UCase(RProveedores(6))
                        RProveedores2(7) = UCase(RProveedores(7))
                        RProveedores2(8) = RProveedores(8)
                    RProveedores2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Proveedores"
                        Err.Clear
                    End If
                RProveedores.MoveNext
            Loop
            
            Exit Sub

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RRutinas = New ADODB.Recordset
            Call Abrir_Recordset(RRutinas, "Select * From Rutinas")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RRutinas2 = New ADODB.Recordset
            Call Abrir_Recordset2(RRutinas2, "Select * From Rutinas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RRutinas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RRutinas2.AddNew
                        RRutinas2(0) = UCase(RRutinas(0))
                        RRutinas2(1) = LCase(RRutinas(1))
                        RRutinas2(2) = RRutinas(2)
                        RRutinas2(3) = RRutinas(3)
                        RRutinas2(4) = RRutinas(4)
                        RRutinas2(5) = RRutinas(5)
                        RRutinas2(6) = UCase(RRutinas(6))
                    RRutinas2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Rutinas"
                        Err.Clear
                    End If
                RRutinas.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RTiposEntradasMP = New ADODB.Recordset
            'Call Abrir_Recordset(RTiposEntradasMP, "Select * From TiposEntradasMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RTiposEntradasMP2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RTiposEntradasMP2, "Select * From TiposEntradasInventario")
            ''HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RTiposEntradasMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RTiposEntradasMP2.AddNew
            '            RTiposEntradasMP2(0) = UCase(RTiposEntradasMP(0))
            '            RTiposEntradasMP2(1) = UCase(RTiposEntradasMP(1))
            '            RTiposEntradasMP2(2) = UCase(RTiposEntradasMP(2))
            '        RTiposEntradasMP2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Tipos Entradas Inventario"
            '            Err.Clear
            '        End If
            '    RTiposEntradasMP.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RTransportistas = New ADODB.Recordset
            'Call Abrir_Recordset(RTransportistas, "Select * From Transportistas")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RTransportistas2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RTransportistas2, "Select * From Transportistas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RTransportistas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RTransportistas2.AddNew
            '            RTransportistas2(0) = UCase(RTransportistas(0))
            '            RTransportistas2(1) = UCase(RTransportistas(1))
            '            RTransportistas2(2) = UCase(RTransportistas(2))
            '            RTransportistas2(3) = UCase(RTransportistas(3))
            '            RTransportistas2(4) = UCase(RTransportistas(4))
            '            RTransportistas2(5) = UCase(RTransportistas(5))
            '            RTransportistas2(6) = UCase(RTransportistas(6))
            '        RTransportistas2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Transportistas"
            '            Err.Clear
            '        End If
            '    RTransportistas.MoveNext
            'Loop


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RUnidadMedida = New ADODB.Recordset
            'Call Abrir_Recordset(RUnidadMedida, "Select * From UnidadesMedida")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RUnidadMedida2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RUnidadMedida2, "Select * From UnidadesMedida")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RUnidadMedida.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RUnidadMedida2.AddNew
            '            RUnidadMedida2(0) = UCase(RUnidadMedida(0))
            '            RUnidadMedida2(1) = UCase(RUnidadMedida(1))
            '            RUnidadMedida2(2) = UCase(RUnidadMedida(2))
            '        RUnidadMedida2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Unidad Medida"
            '            Err.Clear
            '        End If
            '    RUnidadMedida.MoveNext
            'Loop


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RVariablesMedia = New ADODB.Recordset
            Call Abrir_Recordset(RVariablesMedia, "Select * From VariablesMedia")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RVariablesMedia2 = New ADODB.Recordset
            Call Abrir_Recordset2(RVariablesMedia2, "Select * From VariablesMedia")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RVariablesMedia.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RVariablesMedia2.AddNew
                        RVariablesMedia2(0) = UCase(RVariablesMedia(0))
                        RVariablesMedia2(1) = RVariablesMedia(1)
                        RVariablesMedia2(2) = RVariablesMedia(2)
                        RVariablesMedia2(3) = RVariablesMedia(3)
                        RVariablesMedia2(4) = RVariablesMedia(4)
                        RVariablesMedia2(5) = UCase(RVariablesMedia(5))
                    RVariablesMedia2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Variables Media"
                        Err.Clear
                    End If
                RVariablesMedia.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RVentas = New ADODB.Recordset
            Call Abrir_Recordset(RVentas, "Select * From Ventas")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RVentas2 = New ADODB.Recordset
            Call Abrir_Recordset2(RVentas2, "Select * From Ventas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RVentas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RVentas2.AddNew
                        RVentas2(0) = RVentas(0)
                        RVentas2(1) = UCase(RVentas(1))
                        If RVentas(2) = "200" Then
                            RVentas2(2) = "T14"
                        ElseIf RVentas(2) = "201" Then
                            RVentas2(2) = "T15"
                        ElseIf RVentas(2) = "202" Then
                            RVentas2(2) = "T16"
                        ElseIf RVentas(2) = "203" Then
                            RVentas2(2) = "T17"
                        ElseIf RVentas(2) = "204" Then
                            RVentas2(2) = "T18"
                        ElseIf RVentas(2) = "205" Then
                            RVentas2(2) = "T19"
                        ElseIf RVentas(2) = "206" Then
                            RVentas2(2) = "T20"
                        ElseIf RVentas(2) = "207" Then
                            RVentas2(2) = "T21"
                        ElseIf RVentas(2) = "208" Then
                            RVentas2(2) = "T22"
                        ElseIf RVentas(2) = "209" Then
                            RVentas2(2) = "T23"
                        ElseIf RVentas(2) = "210" Then
                            RVentas2(2) = "T24"
                        ElseIf RVentas(2) = "211" Then
                            RVentas2(2) = "T25"
                        End If
                        
                        RVentas2(3) = RVentas(3)
                        RVentas2(4) = UCase(RVentas(4))
                    RVentas2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Ventas"
                        Err.Clear
                    End If
                RVentas.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RVentasDetalle = New ADODB.Recordset
            'Call Abrir_Recordset(RVentasDetalle, "Select * From VentasDetalle")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RVentasDetalle2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RVentasDetalle2, "Select * From VentasDetalle")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RVentasDetalle.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RVentasDetalle2.AddNew
            '            RVentasDetalle2(0) = RVentasDetalle(0)
            '            RVentasDetalle2(1) = RVentasDetalle(1)
            '            RVentasDetalle2(2) = RVentasDetalle(2)
            '            RVentasDetalle2(3) = RVentasDetalle(3)
            '            RVentasDetalle2(4) = RVentasDetalle(4)
            '            RVentasDetalle2(5) = RVentasDetalle(5)
            '            RVentasDetalle2(6) = UCase(RVentasDetalle(6))
            '            RVentasDetalle2(7) = RVentasDetalle(7)
            '            RVentasDetalle2(8) = RVentasDetalle(8)
            '            RVentasDetalle2(9) = UCase(RVentasDetalle(9))
            '        RVentasDetalle2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Ventas Detalle"
            '            Err.Clear
            '        End If
            '    RVentasDetalle.MoveNext
            'Loop

End Sub

Public Sub seis()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoOrdenProduccion = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoOrdenProduccion, "Select * From EncabezadoOrdenProduccion")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoOrdenProduccion2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoOrdenProduccion2, "Select * From EncabezadoOrdenProduccion")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoOrdenProduccion.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoOrdenProduccion2.AddNew
                        REncabezadoOrdenProduccion2(0) = UCase(REncabezadoOrdenProduccion(0))
                        REncabezadoOrdenProduccion2(1) = UCase(REncabezadoOrdenProduccion(1))
                        REncabezadoOrdenProduccion2(2) = REncabezadoOrdenProduccion(2)
                        REncabezadoOrdenProduccion2(3) = REncabezadoOrdenProduccion(3)
                        REncabezadoOrdenProduccion2(4) = UCase(REncabezadoOrdenProduccion(4))
                        REncabezadoOrdenProduccion2(5) = UCase(REncabezadoOrdenProduccion(5))
                        REncabezadoOrdenProduccion2(6) = UCase(REncabezadoOrdenProduccion(6))
                    REncabezadoOrdenProduccion2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Orden Prouduccion"
                        Err.Clear
                    End If
                REncabezadoOrdenProduccion.MoveNext
            Loop
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            
            Set RDetalleOrdenProduccion = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleOrdenProduccion, "Select * From DetalleOrdenProduccion")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleOrdenProduccion2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleOrdenProduccion2, "Select * From DetalleOrdenProduccion")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleOrdenProduccion.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleOrdenProduccion2.AddNew
                        RDetalleOrdenProduccion2(0) = UCase(RDetalleOrdenProduccion(0))
                        RDetalleOrdenProduccion2(1) = UCase(RDetalleOrdenProduccion(1))
                        RDetalleOrdenProduccion2(2) = UCase(RDetalleOrdenProduccion(2))
                        RDetalleOrdenProduccion2(3) = RDetalleOrdenProduccion(3)
                        RDetalleOrdenProduccion2(4) = RDetalleOrdenProduccion(4)
                        RDetalleOrdenProduccion2(5) = RDetalleOrdenProduccion(5)
                        RDetalleOrdenProduccion2(6) = RDetalleOrdenProduccion(6)
                        RDetalleOrdenProduccion2(7) = UCase(RDetalleOrdenProduccion(7))
                    RDetalleOrdenProduccion2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Orden Produccion"
                        Err.Clear
                    End If
                RDetalleOrdenProduccion.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoCapturaParos = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoCapturaParos, "Select * From EncabezadoCapturaParos where year(fecha) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoCapturaParos2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoCapturaParos2, "Select * From EncabezadoCapturaParos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoCapturaParos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoCapturaParos2.AddNew
                        REncabezadoCapturaParos2(0) = REncabezadoCapturaParos(0)
                        REncabezadoCapturaParos2(1) = REncabezadoCapturaParos(1)
                        REncabezadoCapturaParos2(2) = UCase(REncabezadoCapturaParos(2))
                        REncabezadoCapturaParos2(3) = UCase(REncabezadoCapturaParos(3))
                        REncabezadoCapturaParos2(4) = UCase(REncabezadoCapturaParos(4))
                        REncabezadoCapturaParos2(5) = UCase(REncabezadoCapturaParos(5))
                        REncabezadoCapturaParos2(6) = UCase(REncabezadoCapturaParos(6))
                        REncabezadoCapturaParos2(7) = REncabezadoCapturaParos(7)
                        REncabezadoCapturaParos2(8) = REncabezadoCapturaParos(8)
                        REncabezadoCapturaParos2(9) = REncabezadoCapturaParos(9)
                        REncabezadoCapturaParos2(10) = REncabezadoCapturaParos(10)
                        REncabezadoCapturaParos2(11) = REncabezadoCapturaParos(11)
                        REncabezadoCapturaParos2(12) = REncabezadoCapturaParos(12)
                        REncabezadoCapturaParos2(13) = REncabezadoCapturaParos(13)
                        If IsNull(REncabezadoCapturaParos(14)) Then
                            REncabezadoCapturaParos2(14) = 0
                        Else
                            REncabezadoCapturaParos2(14) = REncabezadoCapturaParos(14)
                        End If
                        If IsNull(REncabezadoCapturaParos(15)) Then
                            REncabezadoCapturaParos2(15) = 0
                        Else
                            REncabezadoCapturaParos2(15) = REncabezadoCapturaParos(15)
                        End If
                        REncabezadoCapturaParos2(16) = UCase(REncabezadoCapturaParos(16))
                    REncabezadoCapturaParos2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Captura Paros"
                        Err.Clear
                    End If
                REncabezadoCapturaParos.MoveNext
            Loop
            
            
            MsgBox "ya"
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            
            
            
            
            
            Set RDetalleCapturaParos = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleCapturaParos, "Select * From DetalleCapturaParos where Documento >= 7001 and documento <= 8215")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleCapturaParos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleCapturaParos2, "Select * From DetalleCapturaParos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleCapturaParos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleCapturaParos2.AddNew
                        RDetalleCapturaParos2(0) = RDetalleCapturaParos(0)
                        RDetalleCapturaParos2(1) = UCase(RDetalleCapturaParos(1))
                        RDetalleCapturaParos2(2) = UCase(RDetalleCapturaParos(2))
                        RDetalleCapturaParos2(3) = UCase(RDetalleCapturaParos(3))
                        RDetalleCapturaParos2(4) = RDetalleCapturaParos(4)
                        RDetalleCapturaParos2(5) = UCase(RDetalleCapturaParos(5))
                        If IsNull(RDetalleCapturaParos(6)) Then
                            RDetalleCapturaParos2(6) = "0129"
                        Else
                            RDetalleCapturaParos2(6) = UCase(RDetalleCapturaParos(6))
                        End If
                    RDetalleCapturaParos2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Captura Paros"
                        Err.Clear
                    End If
                RDetalleCapturaParos.MoveNext
            Loop


            MsgBox "·ya"

            
            '___________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleConsumoMP = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleConsumoMP, "Select * From DetalleConsumoMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleConsumoMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleConsumoMP2, "Select * From DetalleConsumoMateriaPrima")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleConsumoMP.EOF
                        
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima as E, DetalleEntradasMateriaPrima as D Where E.Documento = D.Documento And D.Codigo = '" & RDetalleConsumoMP!CodigoMateriaPrima & "' And D.NumeroIngreso = " & RDetalleConsumoMP!NumeroIngreso)
                            If RBuscaFechaEntrada.RecordCount > 0 Then
            
                                    'AGREGA UN REGISTRO EN ORACLE
                                    RDetalleConsumoMP2.AddNew
                                        RDetalleConsumoMP2!Documento = RDetalleConsumoMP!Documento
                                        RDetalleConsumoMP2!Orden = UCase(RDetalleConsumoMP!Orden)
                                        RDetalleConsumoMP2!fecha = RBuscaFechaEntrada(0)
                                        RDetalleConsumoMP2!Linea = "77"
                                        RDetalleConsumoMP2!FichaTecnica = RDetalleConsumoMP!CodigoMateriaPrima
                                        RDetalleConsumoMP2!Tarima = RDetalleConsumoMP!NumeroIngreso
                                        RDetalleConsumoMP2!Desperdicio = UCase(RDetalleConsumoMP!Desperdicio)
                                        RDetalleConsumoMP2!Cantidad = UCase(RDetalleConsumoMP!Cantidad)
                                    RDetalleConsumoMP2.Update
                                    If Err <> 0 Then
                                        MsgBox Err.Description & "Detalle Consumo Materia Prima"
                                        Err.Clear
                                    End If
                                    
                                    
                            Else
                                'MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            End If
                    
                RDetalleConsumoMP.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleProduccionPorOrden = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleProduccionPorOrden, "Select * From DetalleProduccionPorOrden where documento >= 6001 and documento <= 9000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleProduccionPorOrden2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleProduccionPorOrden2, "Select * From DetalleProduccionPorOrden")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleProduccionPorOrden.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleProduccionPorOrden2.AddNew
                        RDetalleProduccionPorOrden2(0) = RDetalleProduccionPorOrden(0)
                        RDetalleProduccionPorOrden2(1) = UCase(RDetalleProduccionPorOrden(1))
                        RDetalleProduccionPorOrden2(2) = RDetalleProduccionPorOrden(2)
                        RDetalleProduccionPorOrden2(3) = RDetalleProduccionPorOrden(3)
                        RDetalleProduccionPorOrden2(4) = RDetalleProduccionPorOrden(4)
                        RDetalleProduccionPorOrden2(5) = RDetalleProduccionPorOrden(5)
                        RDetalleProduccionPorOrden2(6) = UCase(RDetalleProduccionPorOrden(6))
                    RDetalleProduccionPorOrden2.Update
                    If Err = -2147217873 Then
                    ElseIf Err <> 0 Then
                        MsgBox Err.Description & "Detalle Produccion Por Orden"
                        Err.Clear
                    End If
                RDetalleProduccionPorOrden.MoveNext
            Loop

            MsgBox "ya"
End Sub

Public Sub Tres2()

On Error Resume Next
    'Dim cont As Integer
    '        cont = 7102
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoEntradasMP = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoEntradasMP, "Select * From EncabezadoEntradasMateriaPrima Order By Documento")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoEntradasMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoEntradasMP2, "Select * From EncabezadoEntradasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoEntradasMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoEntradasMP2.AddNew
                        REncabezadoEntradasMP2!FechaEntrada = REncabezadoEntradasMP!FechaEntrada
                        REncabezadoEntradasMP2!Documento = REncabezadoEntradasMP!Documento
                        REncabezadoEntradasMP2!Bodega = UCase(REncabezadoEntradasMP!Bodega)
                        REncabezadoEntradasMP2!Batch = "0"
                        REncabezadoEntradasMP2!Linea = "77"
                        REncabezadoEntradasMP2!Observaciones = Mid(REncabezadoEntradasMP!Observaciones, 1, 50)
                        REncabezadoEntradasMP2!ProduccionInterna = "0"
                        REncabezadoEntradasMP2!ProduccionLiberada = "0"
                        If IsNull(REncabezadoEntradasMP!TipoEntrada) Then
                            REncabezadoEntradasMP2!TipoEntrada = "5"
                        Else
                            REncabezadoEntradasMP2!TipoEntrada = UCase(REncabezadoEntradasMP!TipoEntrada)
                        End If
                        If IsNull(REncabezadoEntradasMP!Proveedor) Then
                            REncabezadoEntradasMP2!Proveedor = "40"
                        Else
                            REncabezadoEntradasMP2!Proveedor = UCase(REncabezadoEntradasMP!Proveedor)
                        End If
                        REncabezadoEntradasMP2!NumeroDocumento = UCase(REncabezadoEntradasMP!NumeroDocumento)
                        If IsNull(REncabezadoEntradasMP!TipoDeDocumento) Then
                            REncabezadoEntradasMP2!TipoDeDocumento = "XX"
                        Else
                            REncabezadoEntradasMP2!TipoDeDocumento = UCase(REncabezadoEntradasMP!TipoDeDocumento)
                        End If
                        'NO INGRESABAN CODIGO DETRANSPORTISTA ENTONCES SE LE ASIGNO VARIOS
                        REncabezadoEntradasMP2!Transportista = "VARIOS"
                        REncabezadoEntradasMP2!NombreDePiloto = UCase(REncabezadoEntradasMP!NombreDePiloto)
                        REncabezadoEntradasMP2!PlacasCamion = UCase(REncabezadoEntradasMP!PlacasCamion)
                        REncabezadoEntradasMP2!PlacasFurgon = REncabezadoEntradasMP!PlacasFurgon
                        REncabezadoEntradasMP2!Requerido = UCase(REncabezadoEntradasMP!Requerido)
                        REncabezadoEntradasMP2!Liberado = UCase(REncabezadoEntradasMP!Liberado)
                        REncabezadoEntradasMP2!Estado = UCase(REncabezadoEntradasMP!Estado)
                    REncabezadoEntradasMP2.Update
                    If Err = -2147217887 Then
                        MsgBox Err.Description & "Encabezado Entradas Materia Prima"
                        Err.Clear
                    ElseIf Err <> -2147217887 And Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Entradas Materia Prima"
                        Err.Clear
                        
                        
                         'REncabezadoEntradasMP.MovePrevious
                    End If
                    'cont = cont + 1
                REncabezadoEntradasMP.MoveNext
            Loop


        
            'cont = 7102
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleEntradasMP = New ADODB.Recordset
           Call Abrir_Recordset(RDetalleEntradasMP, "Select * From DetalleEntradasMateriaPrima where documento >= 2601 and documento <= 3000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleEntradasMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleEntradasMP2, "Select * From DetalleEntradasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleEntradasMP.EOF
                    
                    'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima E, DetalleEntradasMateriaPrima D Where E.Documento = D.Documento And D.Codigo = '" & RDetalleEntradasMP!Codigo & "' And D.NumeroIngreso = " & RDetalleEntradasMP!NumeroIngreso)
                            If RBuscaFechaEntrada.RecordCount > 0 Then
                                
                            Else
                                MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            End If
                        
            
            
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleEntradasMP2.AddNew
                        RDetalleEntradasMP2!Documento = RDetalleEntradasMP!Documento
                        RDetalleEntradasMP2!Fechaproduccion = RBuscaFechaEntrada(0)
                        RDetalleEntradasMP2!Linea = "77"
                        RDetalleEntradasMP2!FichaTecnica = UCase(RDetalleEntradasMP!Codigo)
                        RDetalleEntradasMP2!Tarima = RDetalleEntradasMP!NumeroIngreso
                        RDetalleEntradasMP2!Batch = "0"
                        RDetalleEntradasMP2!Calidad = "A"
                        RDetalleEntradasMP2!Bodega = UCase(RDetalleEntradasMP!BodegaDisponibilidad)
                        RDetalleEntradasMP2!Pasillo = RDetalleEntradasMP!Pasillo
                        RDetalleEntradasMP2!Casilla = RDetalleEntradasMP!Casilla
                        RDetalleEntradasMP2!Bin = RDetalleEntradasMP!Bin
                        RDetalleEntradasMP2!Saldo = RDetalleEntradasMP!SaldoDisponibilidad
                        RDetalleEntradasMP2!OrdenProduccion = UCase(RDetalleEntradasMP!OrdenProduccion)
                        RDetalleEntradasMP2!Barra = Format(RBuscaFechaEntrada(0), "ddmmyy") & "77" & UCase(RDetalleEntradasMP!Codigo) & RDetalleEntradasMP!NumeroIngreso
                        RDetalleEntradasMP2!PesoEntrada = RDetalleEntradasMP!PesoEntrada
                        RDetalleEntradasMP2!Estado = "INSPECCIONADO"
                        RDetalleEntradasMP2!SerieBoleta = RDetalleEntradasMP!NumeroUnicoSerieBoleta
                        RDetalleEntradasMP2!OrdenBoleta = RDetalleEntradasMP!OrdenBoleta
                        RDetalleEntradasMP2!BultoBoleta = RDetalleEntradasMP!BultoBoleta
                        RDetalleEntradasMP2!FechaBoleta = RDetalleEntradasMP!FechaBoleta
                        RDetalleEntradasMP2!BobinaBoleta = RDetalleEntradasMP!BobinaBoleta
                        RDetalleEntradasMP2!CantidadEntrada = RDetalleEntradasMP!Cantidad
                        RDetalleEntradasMP2!Observaciones = RDetalleEntradasMP!Observaciones
                    RDetalleEntradasMP2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Entradas Materia Prima"
                        Err.Clear
                    End If
                RDetalleEntradasMP.MoveNext
            Loop
            
            
            MsgBox "ya"
    
            Exit Sub
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoEntradasPT = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoEntradasPT, "Select * From EncabezadoEntradasProductoTerminado") ' where Documento > 11000 And Documento <= 12000")
            
            'Cont = 2793
            '        Do Until REncabezadoEntradasPT.EOF
            '                    REncabezadoEntradasPT!Documento = Cont
            '                REncabezadoEntradasPT.Update
            '               If Err <> 0 Then
            '                    MsgBox Err.Number & Err.Description
            '                    Err.Clear
            '                End If
            '            REncabezadoEntradasPT.MoveNext
            '            Cont = Cont + 1
            '        Loop
           '
            '        MsgBox "ya"
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REncabezadoEntradasPT2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REncabezadoEntradasPT2, "Select * From EncabezadoEntradasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoEntradasPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
             '       REncabezadoEntradasPT2.AddNew
             '           REncabezadoEntradasPT2!FechaEntrada = REncabezadoEntradasPT!FechaEntrada
             '           REncabezadoEntradasPT2!Documento = REncabezadoEntradasPT!Documento
             '           REncabezadoEntradasPT2!Bodega = UCase(REncabezadoEntradasPT!Bodega)
             '           REncabezadoEntradasPT2!Batch = "0"
             '           REncabezadoEntradasPT2!Linea = "77"
             '           REncabezadoEntradasPT2!Observaciones = REncabezadoEntradasPT!Observaciones
             '           REncabezadoEntradasPT2!ProduccionInterna = "0"
             '           REncabezadoEntradasPT2!ProduccionLiberada = "0"
             '           REncabezadoEntradasPT2!TipoEntrada = "5"
             '           REncabezadoEntradasPT2!Proveedor = "VARIOS"
             '           REncabezadoEntradasPT2!NumeroDocumento = "0"
             '           REncabezadoEntradasPT2!TipoDeDocumento = "XX"
             '           REncabezadoEntradasPT2!Transportista = "VARIOS"
             '           REncabezadoEntradasPT2!NombreDePiloto = ""
             '           REncabezadoEntradasPT2!PlacasCamion = ""
             '           REncabezadoEntradasPT2!PlacasFurgon = ""
             '           REncabezadoEntradasPT2!Requerido = UCase(REncabezadoEntradasPT!Requerido)
             '           REncabezadoEntradasPT2!Liberado = UCase(REncabezadoEntradasPT!Liberado)
             '           REncabezadoEntradasPT2!Estado = UCase(REncabezadoEntradasPT!Estado)
             '       REncabezadoEntradasPT2.Update
                
                    VProduccionInterna = REncabezadoEntradasPT!ProduccionInterna
                    VProduccionLiberada = REncabezadoEntradasPT!ProduccionLiberada
                    
                    Conexion2.Execute "Update EncabezadoEntradasInventario Set Batch = " & REncabezadoEntradasPT!Batch & ", Linea = '" & REncabezadoEntradasPT!Linea & "', ProduccionInterna = " & VProduccionInterna & ", ProduccionLiberada = " & VProduccionLiberada & " Where Documento = " & REncabezadoEntradasPT!Documento
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Entradas Producto Terminado"
                        Err.Clear
                    End If
                REncabezadoEntradasPT.MoveNext
            Loop



MsgBox "ya"
Exit Sub

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleEntradasPT = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleEntradasPT, "Select * From DetalleEntradasProductoTerminaDO where Documento >= 5001 and Documento <= 5500")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleEntradasPT2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleEntradasPT2, "Select * From DetalleEntradasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleEntradasPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    
                    RDetalleEntradasPT2.AddNew
                        RDetalleEntradasPT2!Documento = RDetalleEntradasPT!Documento
                        RDetalleEntradasPT2!Fechaproduccion = RDetalleEntradasPT!Fechaproduccion
                        RDetalleEntradasPT2!Linea = UCase(RDetalleEntradasPT!Linea)
                        RDetalleEntradasPT2!FichaTecnica = UCase(RDetalleEntradasPT!FichaTecnica)
                        RDetalleEntradasPT2!Tarima = RDetalleEntradasPT!Tarima
                        RDetalleEntradasPT2!Batch = RDetalleEntradasPT!Batch
                        RDetalleEntradasPT2!Calidad = RDetalleEntradasPT!Calidad
                        RDetalleEntradasPT2!Bodega = UCase(RDetalleEntradasPT!Bodega)
                        RDetalleEntradasPT2!Pasillo = RDetalleEntradasPT!Pasillo
                        RDetalleEntradasPT2!Casilla = RDetalleEntradasPT!Casilla
                        RDetalleEntradasPT2!Bin = RDetalleEntradasPT!Bin
                        RDetalleEntradasPT2!Saldo = RDetalleEntradasPT!Saldo
                        RDetalleEntradasPT2!OrdenProduccion = UCase(RDetalleEntradasPT!OrdenProduccion)
                        RDetalleEntradasPT2!Barra = Format(RDetalleEntradasPT!Fechaproduccion, "ddmmyy") & RDetalleEntradasPT!Linea & RDetalleEntradasPT!FichaTecnica & RDetalleEntradasPT!Tarima
                        RDetalleEntradasPT2!PesoEntrada = "0"
                        RDetalleEntradasPT2!Estado = "INSPECCIONADO"
                        RDetalleEntradasPT2!SerieBoleta = ""
                        RDetalleEntradasPT2!OrdenBoleta = ""
                        RDetalleEntradasPT2!BultoBoleta = ""
                        RDetalleEntradasPT2!FechaBoleta = ""
                        RDetalleEntradasPT2!BobinaBoleta = ""
                        RDetalleEntradasPT2!CantidadEntrada = RDetalleEntradasPT!Cantidad
                        RDetalleEntradasPT2!Observaciones = ""
                    RDetalleEntradasPT2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Entradas Producto Terminado"
                        Err.Clear
                    End If
                RDetalleEntradasPT.MoveNext
            Loop
            
            
            MsgBox "ya"

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoDespachosPT = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoDespachosPT, "Select * From EncabezadoDespachosProductoTerminado") ' where Documento > 7000 And Documento <= 8000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoDespachosPT2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoDespachosPT2, "Select * From EncabezadoSalidasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoDespachosPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoDespachosPT2.AddNew
                        REncabezadoDespachosPT2(0) = REncabezadoDespachosPT(0)
                        REncabezadoDespachosPT2(1) = REncabezadoDespachosPT(1)
                        REncabezadoDespachosPT2(2) = UCase(REncabezadoDespachosPT(2))
                        REncabezadoDespachosPT2(3) = REncabezadoDespachosPT(3)
                        REncabezadoDespachosPT2(4) = UCase(REncabezadoDespachosPT(4))
                        REncabezadoDespachosPT2(5) = UCase(REncabezadoDespachosPT(5))
                        REncabezadoDespachosPT2(6) = UCase(REncabezadoDespachosPT(6))
                        REncabezadoDespachosPT2(7) = UCase(REncabezadoDespachosPT(7))
                        REncabezadoDespachosPT2(8) = UCase(REncabezadoDespachosPT(8))
                        REncabezadoDespachosPT2(9) = UCase(REncabezadoDespachosPT(9))
                        REncabezadoDespachosPT2(10) = UCase(REncabezadoDespachosPT(10))
                        REncabezadoDespachosPT2(11) = UCase(REncabezadoDespachosPT(11))
                        REncabezadoDespachosPT2(12) = UCase(REncabezadoDespachosPT(12))
                        REncabezadoDespachosPT2(13) = UCase(REncabezadoDespachosPT(13))
                        REncabezadoDespachosPT2(14) = UCase(REncabezadoDespachosPT(14))
                        REncabezadoDespachosPT2(15) = UCase(REncabezadoDespachosPT(15))
                        REncabezadoDespachosPT2(16) = UCase(REncabezadoDespachosPT(16))
                    REncabezadoDespachosPT2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Despachos Producto Terminado"
                        Err.Clear
                    End If
                REncabezadoDespachosPT.MoveNext
            Loop
            
            
                        
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleDespachosPT = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleDespachosPT, "Select * From DetalleDespachosProductoTerminado where Documento >= 2001 And Documento <= 2500 Order By Documento")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleDespachosPT2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleDespachosPT2, "Select * From DetalleSalidasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleDespachosPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleDespachosPT2.AddNew
                        RDetalleDespachosPT2(0) = RDetalleDespachosPT(0)
                        RDetalleDespachosPT2(1) = RDetalleDespachosPT(1)
                        RDetalleDespachosPT2(2) = UCase(RDetalleDespachosPT(2))
                        RDetalleDespachosPT2(3) = UCase(RDetalleDespachosPT(3))
                        RDetalleDespachosPT2(4) = RDetalleDespachosPT(4)
                        RDetalleDespachosPT2(5) = RDetalleDespachosPT(5)
                        If IsNull(RDetalleDespachosPT(6)) Then
                            RDetalleDespachosPT2(6) = "A"
                        Else
                            RDetalleDespachosPT2(6) = UCase(RDetalleDespachosPT(6))
                        End If
                        RDetalleDespachosPT2(7) = UCase(RDetalleDespachosPT(7))
                        RDetalleDespachosPT2(8) = RDetalleDespachosPT(8)
                    RDetalleDespachosPT2.Update
                    If Err <> 0 And Err <> -2147467259 Then
                        MsgBox RDetalleDespachosPT(1) & " " & RDetalleDespachosPT(4) & " " & Err.Description & "Detalle Despachos Producto Terminado"
                        Err.Clear
                    End If
                RDetalleDespachosPT.MoveNext
            Loop
            
            
            
            
            
            
            'EN ACCESS CAMBIAR EL CORRELATIVO DE DOCUMENTOS
            Set REncabezadoEgresosMP = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoEgresosMP, "Select * From EncabezadoEgresosMateriaPrima ORDER BY Documento")
                    Cont = 2306 'EL ULTIMO DOCUMENTO DE PT
                    Do Until REncabezadoEgresosMP.EOF
                                REncabezadoEgresosMP!Documento = Cont
                                REncabezadoEgresosMP.Update
                                If Err <> 0 Then
                                    MsgBox Err.Description
                                End If
                            REncabezadoEgresosMP.MoveNext
                            Cont = Cont + 1
                    Loop
            
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoEgresosMP = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoEgresosMP, "Select * From EncabezadoEgresosMateriaPrima") ' Where Documento > 9000 And Documento <= 10000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoEgresosMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoEgresosMP2, "Select * From EncabezadoSalidasInventario")
            
            
            
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoEgresosMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoEgresosMP2.AddNew
                        REncabezadoEgresosMP2!Documento = REncabezadoEgresosMP!Documento
                        REncabezadoEgresosMP2!fecha = REncabezadoEgresosMP!fecha
                        REncabezadoEgresosMP2!Cliente = UCase(REncabezadoEgresosMP!Cliente)
                        REncabezadoEgresosMP2!Batch = "0"
                        REncabezadoEgresosMP2!Linea = "77"
                        REncabezadoEgresosMP2!CodigoTransportista = UCase(REncabezadoEgresosMP!CodigoTransportista)
                        REncabezadoEgresosMP2!TipoDeDocumento = UCase(REncabezadoEgresosMP!TipoDeDocumento)
                        REncabezadoEgresosMP2!NumeroDocumento = UCase(REncabezadoEgresosMP!NumeroDocumento)
                        REncabezadoEgresosMP2!CargadoPor = UCase(REncabezadoEgresosMP!CargadoPor)
                        REncabezadoEgresosMP2!EntregadoPor = UCase(REncabezadoEgresosMP!EntregadoPor)
                        REncabezadoEgresosMP2!Conductor = UCase(REncabezadoEgresosMP!Conductor)
                        REncabezadoEgresosMP2!PlacasCamion = UCase(REncabezadoEgresosMP!PlacasCamion)
                        REncabezadoEgresosMP2!PlacasFurgon = UCase(REncabezadoEgresosMP!PlacasFurgon)
                        REncabezadoEgresosMP2!Observaciones = UCase(REncabezadoEgresosMP!Observaciones)
                        REncabezadoEgresosMP2!Requerido = UCase(REncabezadoEgresosMP!Requerido)
                        REncabezadoEgresosMP2!Liberado = UCase(REncabezadoEgresosMP!Liberado)
                        REncabezadoEgresosMP2!Estado = UCase(REncabezadoEgresosMP!Estado)
                    REncabezadoEgresosMP2.Update
                    If Err = 3265 Then
                        MsgBox Err.Description & "Encabezado Egresos Materia Prima"
                        Err.Clear
                    ElseIf Err <> 0 And Err <> 3265 Then
                        MsgBox Err.Description & "Encabezado Egresos Materia Prima"
                        Err.Clear
                    End If
                REncabezadoEgresosMP.MoveNext
                
            Loop
            
            
            MsgBox "ya"
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleEgresosMP = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleEgresosMP, "Select * From DetalleEgresosMateriaPrima Where Documento >= 3501 And Documento <= 4000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleEgresosMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleEgresosMP2, "Select * From DetalleSalidasInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleEgresosMP.EOF
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                           Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima as E, DetalleEntradasMateriaPrima as D Where E.Documento = D.Documento And D.Codigo = '" & RDetalleEgresosMP!Codigo & "' And D.NumeroIngreso = " & RDetalleEgresosMP!NumeroIngreso)
                            If RBuscaFechaEntrada.RecordCount > 0 Then
                                
                            
            
            
                                    'AGREGA UN REGISTRO EN ORACLE
                                    RDetalleEgresosMP2.AddNew
                                        RDetalleEgresosMP2!Documento = RDetalleEgresosMP!Documento
                                        RDetalleEgresosMP2!Fechaproduccion = RBuscaFechaEntrada(0)
                                        RDetalleEgresosMP2!Linea = "77"
                                        RDetalleEgresosMP2!FichaTecnica = UCase(RDetalleEgresosMP!Codigo)
                                        RDetalleEgresosMP2!Tarima = RDetalleEgresosMP!NumeroIngreso
                                        RDetalleEgresosMP2!Batch = "0"
                                        RDetalleEgresosMP2!Calidad = "A"
                                        RDetalleEgresosMP2!Bodega = UCase(RDetalleEgresosMP!Bodega)
                                        RDetalleEgresosMP2!Cantidad = RDetalleEgresosMP!Cantidad
                                    RDetalleEgresosMP2.Update
                                    If Err = -2147217873 Then
                                        MsgBox Err.Description & "Detalle Egresos Materia Prima"
                                        Err.Clear
                                    ElseIf Err <> 0 And Err <> -2147217873 Then
                                        MsgBox Err.Description & "Detalle Egresos Materia Prima"
                                        Err.Clear
                                    End If
                            Else
                                'MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
    
                    
                    End If
                RDetalleEgresosMP.MoveNext
            Loop
            
            MsgBox "ya"


'___________TOMAR EN CUENTA QUE SE TRASLADA POR AÑO Y MES___________________________________________________________________________________________________________
                        
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RNumerosIngresosProcesados = New ADODB.Recordset
            Call Abrir_Recordset(RNumerosIngresosProcesados, "Select * From NumerosIngresosProcesados where Year(Fecha) = 2005")
            
            If RNumerosIngresosProcesados.RecordCount > 0 Then
            
            Else
                MsgBox "NO hay bultos"
            
            End If
            'ABRIMOS EL RECORDSET DE ORACLE
            
            Set RNumerosIngresosProcesados2 = New ADODB.Recordset
            Call Abrir_Recordset2(RNumerosIngresosProcesados2, "Select * From CierreBulto")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RNumerosIngresosProcesados.EOF
                        
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima as E, DetalleEntradasMateriaPrima as D Where E.Documento = D.Documento And D.Codigo = '" & UCase(RNumerosIngresosProcesados!CodigoMateriaPrima) & "' And D.NumeroIngreso = " & RNumerosIngresosProcesados!NumeroIngreso)
                            
                           If RBuscaFechaEntrada.RecordCount > 0 Then
                                                   
                                       
                                        'AGREGA UN REGISTRO EN ORACLE
                                        RNumerosIngresosProcesados2.AddNew
                                            RNumerosIngresosProcesados2!fecha = RNumerosIngresosProcesados!fecha
                                            RNumerosIngresosProcesados2!Turno = RNumerosIngresosProcesados!Turno
                                            RNumerosIngresosProcesados2!Linea = UCase(RNumerosIngresosProcesados!Linea)
                                            RNumerosIngresosProcesados2!BodegaSalida = UCase(RNumerosIngresosProcesados!BodegaSalida)
                                            RNumerosIngresosProcesados2!Existencia = RNumerosIngresosProcesados!Existencia
                                            RNumerosIngresosProcesados2!CantidadMas = RNumerosIngresosProcesados!CantidadMas
                                            RNumerosIngresosProcesados2!CantidadMenos = RNumerosIngresosProcesados!CantidadMenos
                                            RNumerosIngresosProcesados2!ExistenciaNueva = RNumerosIngresosProcesados!ExistenciaNueva
                                            RNumerosIngresosProcesados2!ContadorInicial = RNumerosIngresosProcesados!ContadorInicial
                                            RNumerosIngresosProcesados2!ContadorFinal = RNumerosIngresosProcesados!ContadorFinal
                                            RNumerosIngresosProcesados2!CantidadProcesada = RNumerosIngresosProcesados!CantidadProcesada
                                            RNumerosIngresosProcesados2!DesperdicioProceso = RNumerosIngresosProcesados!DesperdicioProceso
                                            RNumerosIngresosProcesados2!DesperdicioProveedor = RNumerosIngresosProcesados!DesperdicioProveedor
                                            RNumerosIngresosProcesados2!CantidadProcesadaReal = RNumerosIngresosProcesados!CantidadProcesadaReal
                                            RNumerosIngresosProcesados2!Total = RNumerosIngresosProcesados!Total
                                            RNumerosIngresosProcesados2!Usuario = UCase(RNumerosIngresosProcesados!UsuarioAgregar)
                                            RNumerosIngresosProcesados2!Fechaproduccion = RBuscaFechaEntrada(0)
                                            RNumerosIngresosProcesados2!LineaProduccion = "77"
                                            RNumerosIngresosProcesados2!FichaTecnica = UCase(RNumerosIngresosProcesados!CodigoMateriaPrima)
                                            RNumerosIngresosProcesados2!Tarima = RNumerosIngresosProcesados!NumeroIngreso
                                            RNumerosIngresosProcesados2!Hora = Format(RNumerosIngresosProcesados!Hora, "hh:mm")
                                            RNumerosIngresosProcesados2!Observaciones = ""
                                        RNumerosIngresosProcesados2.Update
                            Else
                                'MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                                'Err.Clear
                            End If
                                        
                    If Err <> 0 Then
                        MsgBox Err.Description & "Numeros Ingresos Procesados"
                        Err.Clear
                    End If
                RNumerosIngresosProcesados.MoveNext
            Loop
'______________________________________________________________________________________________________________________

MsgBox "YA"
            
End Sub

Public Sub ocho()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoPedidosProveedores = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoPedidosProveedores, "Select * From EncabezadoPedidosProveedores")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoPedidosProveedores2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoPedidosProveedores2, "Select * From EncabezadoPedidosProveedores")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoPedidosProveedores.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoPedidosProveedores2.AddNew
                        REncabezadoPedidosProveedores2(0) = REncabezadoPedidosProveedores(0)
                        REncabezadoPedidosProveedores2(1) = UCase(REncabezadoPedidosProveedores(1))
                        REncabezadoPedidosProveedores2(2) = UCase(REncabezadoPedidosProveedores(2))
                        REncabezadoPedidosProveedores2(3) = UCase(REncabezadoPedidosProveedores(3))
                        REncabezadoPedidosProveedores2(4) = UCase(REncabezadoPedidosProveedores(4))
                    REncabezadoPedidosProveedores2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Pedidos Proveedores"
                        Err.Clear
                    End If
                REncabezadoPedidosProveedores.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetallePedidosProveedores = New ADODB.Recordset
            Call Abrir_Recordset(RDetallePedidosProveedores, "Select * From DetallePedidosProveedores")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetallePedidosProveedores2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetallePedidosProveedores2, "Select * From DetallePedidosProveedores")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetallePedidosProveedores.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetallePedidosProveedores2.AddNew
                        RDetallePedidosProveedores2(0) = UCase(RDetallePedidosProveedores(0))
                        RDetallePedidosProveedores2(1) = UCase(RDetallePedidosProveedores(1))
                        RDetallePedidosProveedores2(2) = RDetallePedidosProveedores(2)
                        RDetallePedidosProveedores2(3) = RDetallePedidosProveedores(3)
                        RDetallePedidosProveedores2(4) = RDetallePedidosProveedores(4)
                        RDetallePedidosProveedores2(5) = RDetallePedidosProveedores(5)
                        RDetallePedidosProveedores2(6) = RDetallePedidosProveedores(6)
                        RDetallePedidosProveedores2(7) = RDetallePedidosProveedores(7)
                        RDetallePedidosProveedores2(8) = RDetallePedidosProveedores(8)
                    RDetallePedidosProveedores2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Pedidos Proveedores"
                        Err.Clear
                    End If
                RDetallePedidosProveedores.MoveNext
            Loop


'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REncabezadoPedidosClientes = New ADODB.Recordset
            'Call Abrir_Recordset(REncabezadoPedidosClientes, "Select * From EncabezadoPedidosClientes")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REncabezadoPedidosClientes2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REncabezadoPedidosClientes2, "Select * From EncabezadoPedidosClientes")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REncabezadoPedidosClientes.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REncabezadoPedidosClientes2.AddNew
            '            REncabezadoPedidosClientes2(0) = REncabezadoPedidosClientes(0)
            '            REncabezadoPedidosClientes2(1) = UCase(REncabezadoPedidosClientes(1))
            '            REncabezadoPedidosClientes2(2) = UCase(REncabezadoPedidosClientes(2))
            '            REncabezadoPedidosClientes2(3) = UCase(REncabezadoPedidosClientes(3))
            '            REncabezadoPedidosClientes2(4) = UCase(REncabezadoPedidosClientes(4))
            '        REncabezadoPedidosClientes2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Encabezado Pedidos Clientes"
            '            Err.Clear
            '        End If
            '    REncabezadoPedidosClientes.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RDetallePedidosClientes = New ADODB.Recordset
            'Call Abrir_Recordset(RDetallePedidosClientes, "Select * From DetallePedidosClientes")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RDetallePedidosClientes2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RDetallePedidosClientes2, "Select * From DetallePedidosClientes")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RDetallePedidosClientes.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RDetallePedidosClientes2.AddNew
            '            RDetallePedidosClientes2(0) = UCase(RDetallePedidosClientes(0))
            '            RDetallePedidosClientes2(1) = UCase(RDetallePedidosClientes(1))
            '            RDetallePedidosClientes2(2) = RDetallePedidosClientes(2)
            '            RDetallePedidosClientes2(3) = RDetallePedidosClientes(3)
            '            RDetallePedidosClientes2(4) = RDetallePedidosClientes(4)
            '            RDetallePedidosClientes2(5) = RDetallePedidosClientes(5)
            '            RDetallePedidosClientes2(6) = RDetallePedidosClientes(6)
            '            RDetallePedidosClientes2(7) = RDetallePedidosClientes(7)
            '            RDetallePedidosClientes2(8) = RDetallePedidosClientes(8)
            '        RDetallePedidosClientes2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Detalle Pedidos Clientes"
            '            Err.Clear
            '        End If
            '    RDetallePedidosClientes.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoCierrePedidosProveedores = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoCierrePedidosProveedores, "Select * From EncabezadoCierrePedidosProveedores")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoCierrePedidosProveedores2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoCierrePedidosProveedores2, "Select * From EncabezadoCierrePedidosProve")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoCierrePedidosProveedores.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoCierrePedidosProveedores2.AddNew
                        REncabezadoCierrePedidosProveedores2(0) = REncabezadoCierrePedidosProveedores(0)
                        REncabezadoCierrePedidosProveedores2(1) = REncabezadoCierrePedidosProveedores(2)
                        REncabezadoCierrePedidosProveedores2(2) = REncabezadoCierrePedidosProveedores(3)
                        REncabezadoCierrePedidosProveedores2(3) = UCase(REncabezadoCierrePedidosProveedores(4))
                        REncabezadoCierrePedidosProveedores2(4) = UCase(REncabezadoCierrePedidosProveedores(5))
                        REncabezadoCierrePedidosProveedores2(5) = UCase(REncabezadoCierrePedidosProveedores(6))
                    REncabezadoCierrePedidosProveedores2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Cierre Pedidos Proveedores"
                        Err.Clear
                    End If
                REncabezadoCierrePedidosProveedores.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleCierrePedidosProveedores = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleCierrePedidosProveedores, "Select * From DetalleCierrePedidosProveedores")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleCierrePedidosProveedores2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleCierrePedidosProveedores2, "Select * From DetalleCierrePedidosProve")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleCierrePedidosProveedores.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleCierrePedidosProveedores2.AddNew
                        RDetalleCierrePedidosProveedores2(0) = RDetalleCierrePedidosProveedores(0)
                        RDetalleCierrePedidosProveedores2(1) = UCase(RDetalleCierrePedidosProveedores(1))
                        RDetalleCierrePedidosProveedores2(2) = UCase(RDetalleCierrePedidosProveedores(2))
                        RDetalleCierrePedidosProveedores2(3) = RDetalleCierrePedidosProveedores(3)
                    RDetalleCierrePedidosProveedores2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Detalle Cierre Pedidos Proveedores"
                        Err.Clear
                    End If
                RDetalleCierrePedidosProveedores.MoveNext
            Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set REncabezadoCierrePedidosClientes = New ADODB.Recordset
            'Call Abrir_Recordset(REncabezadoCierrePedidosClientes, "Select * From EncabezadoCierrePedidosClientes")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set REncabezadoCierrePedidosClientes2 = New ADODB.Recordset
            'Call Abrir_Recordset2(REncabezadoCierrePedidosClientes2, "Select * From EncabezadoCierrePedidosCliente")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until REncabezadoCierrePedidosClientes.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        REncabezadoCierrePedidosClientes2.AddNew
            '            REncabezadoCierrePedidosClientes2(0) = REncabezadoCierrePedidosClientes(0)
            '            REncabezadoCierrePedidosClientes2(1) = REncabezadoCierrePedidosClientes(1)
            '            REncabezadoCierrePedidosClientes2(2) = REncabezadoCierrePedidosClientes(2)
            '            REncabezadoCierrePedidosClientes2(3) = UCase(REncabezadoCierrePedidosClientes(3))
            '            REncabezadoCierrePedidosClientes2(4) = UCase(REncabezadoCierrePedidosClientes(4))
            '            REncabezadoCierrePedidosClientes2(5) = UCase(REncabezadoCierrePedidosClientes(5))
            '
            '        REncabezadoCierrePedidosClientes2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Encabezado Cierre Pedidos Clientes"
            '            Err.Clear
            '        End If
            '    REncabezadoCierrePedidosClientes.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RDetalleCierrePedidosClientes = New ADODB.Recordset
            'Call Abrir_Recordset(RDetalleCierrePedidosClientes, "Select * From DetalleCierrePedidosClientes")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RDetalleCierrePedidosClientes2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RDetalleCierrePedidosClientes2, "Select * From DetalleCierrePedidosCliente")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RDetalleCierrePedidosClientes.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RDetalleCierrePedidosClientes2.AddNew
            '            RDetalleCierrePedidosClientes2(0) = RDetalleCierrePedidosClientes(0)
            '            RDetalleCierrePedidosClientes2(1) = RDetalleCierrePedidosClientes(1)
            '            RDetalleCierrePedidosClientes2(2) = UCase(RDetalleCierrePedidosClientes(2))
            '            RDetalleCierrePedidosClientes2(3) = UCase(RDetalleCierrePedidosClientes(3))
            '        RDetalleCierrePedidosClientes2.Update
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Detalle Cierre Pedidos Clientes"
            '            Err.Clear
            '        End If
            '    RDetalleCierrePedidosClientes.MoveNext
            'Loop

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoTrasladosMP = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoTrasladosMP, "Select * From EncabezadoTrasladosMateriaPrimaP where year(fecha) >= 2003")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoTrasladosMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoTrasladosMP2, "Select * From EncabezadoTrasladosInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoTrasladosMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoTrasladosMP2.AddNew
                        REncabezadoTrasladosMP2!Documento = REncabezadoTrasladosMP!Documento
                        REncabezadoTrasladosMP2!fecha = REncabezadoTrasladosMP!fecha
                        REncabezadoTrasladosMP2!TipoDeDocumento = UCase(REncabezadoTrasladosMP!TipoDeDocumento)
                        REncabezadoTrasladosMP2!NumeroDocumento = UCase(REncabezadoTrasladosMP!NumeroDocumento)
                        REncabezadoTrasladosMP2!BodegaSalida = UCase(REncabezadoTrasladosMP!BodegaSalida)
                        REncabezadoTrasladosMP2!Requerido = UCase(REncabezadoTrasladosMP!Requerido)
                        REncabezadoTrasladosMP2!Liberado = UCase(REncabezadoTrasladosMP!Liberado)
                        REncabezadoTrasladosMP2!Observaciones = UCase(REncabezadoTrasladosMP!Observaciones)
                        REncabezadoTrasladosMP2!Estado = UCase(REncabezadoTrasladosMP!Estado)
                    REncabezadoTrasladosMP2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Traslados Materia Prima"
                        Err.Clear
                    End If
                REncabezadoTrasladosMP.MoveNext
            Loop

                            
                MsgBox "ya  "
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleTrasladosMP = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleTrasladosMP, "Select * From DetalleTrasladosMateriaPrimaP where Documento >= 5001 and Documento <= 6100")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleTrasladosMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleTrasladosMP2, "Select * From DetalleTrasladosInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleTrasladosMP.EOF
            
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima as E, DetalleEntradasMateriaPrima as D Where E.Documento = D.Documento And D.Codigo = '" & RDetalleTrasladosMP!CodigoSalida & "' And D.NumeroIngreso = " & RDetalleTrasladosMP!NumeroIngreso)
                            
                            If RBuscaFechaEntrada.RecordCount > 0 Then
                                
                            
                                        'AGREGA UN REGISTRO EN ORACLE
                                        RDetalleTrasladosMP2.AddNew
                                            RDetalleTrasladosMP2!Documento = RDetalleTrasladosMP!Documento
                                            RDetalleTrasladosMP2!Tarima = RDetalleTrasladosMP!NumeroIngreso
                                            RDetalleTrasladosMP2!FichaTecnica = UCase(RDetalleTrasladosMP!CodigoSalida)
                                            RDetalleTrasladosMP2!CantidadSalida = RDetalleTrasladosMP!CantidadSalida
                                            RDetalleTrasladosMP2!BodegaEntrada = UCase(RDetalleTrasladosMP!BodegaEntrada)
                                            RDetalleTrasladosMP2!DiferenciaReqCorMas = RDetalleTrasladosMP!DiferenciaReqCorMas
                                            RDetalleTrasladosMP2!DiferenciaReqCor = RDetalleTrasladosMP!DiferenciaReqCor
                                            RDetalleTrasladosMP2!CantidadDesperdicio = RDetalleTrasladosMP!CantidadDesperdicio
                                            RDetalleTrasladosMP2!CantidadDesperdicioProveedor = RDetalleTrasladosMP!CantidadDesperdicioProveedor
                                            RDetalleTrasladosMP2!CantidadReal = RDetalleTrasladosMP!CantidadReal
                                            RDetalleTrasladosMP2!Orden = UCase(RDetalleTrasladosMP!Orden)
                                            RDetalleTrasladosMP2!Fechaproduccion = RBuscaFechaEntrada(0)
                                            RDetalleTrasladosMP2!LineaProduccion = "77"
                                        RDetalleTrasladosMP2.Update
                                        If Err <> 0 Then
                                            MsgBox Err.Description & "Detalle Traslados Materia Prima"
                                            Err.Clear
                                        End If
                            Else
                                'MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            End If
                RDetalleTrasladosMP.MoveNext
            Loop

        MsgBox "ya"
'______________________________________________________________________________________________________________________
            
            
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoTrasladosPT = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoTrasladosPT, "Select * From EncabezadoTrasladosProductoTerminado")
                Cont = 6073
                    Do Until REncabezadoTrasladosPT.EOF
                                REncabezadoTrasladosPT!Documento = Cont
                            REncabezadoTrasladosPT.Update
                           If Err <> 0 Then
                                MsgBox Err.Number & Err.Description
                                Err.Clear
                            End If
                        REncabezadoTrasladosPT.MoveNext
                        Cont = Cont + 1
                    Loop
            
                    MsgBox "ya"
            
            
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REncabezadoTrasladosPT = New ADODB.Recordset
            Call Abrir_Recordset(REncabezadoTrasladosPT, "Select * From EncabezadoTrasladosProductoTerminado")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REncabezadoTrasladosPT2 = New ADODB.Recordset
            Call Abrir_Recordset2(REncabezadoTrasladosPT2, "Select * From EncabezadoTrasladosInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REncabezadoTrasladosPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REncabezadoTrasladosPT2.AddNew
                        REncabezadoTrasladosPT2!Documento = REncabezadoTrasladosPT!Documento
                        REncabezadoTrasladosPT2!fecha = REncabezadoTrasladosPT!fecha
                        REncabezadoTrasladosPT2!TipoDeDocumento = UCase(REncabezadoTrasladosPT!TipoDeDocumento)
                        REncabezadoTrasladosPT2!NumeroDocumento = UCase(REncabezadoTrasladosPT!NumeroDocumento)
                        REncabezadoTrasladosPT2!BodegaSalida = "001"
                        REncabezadoTrasladosPT2!Requerido = UCase(REncabezadoTrasladosPT!Requerido)
                        REncabezadoTrasladosPT2!Liberado = UCase(REncabezadoTrasladosPT!Liberado)
                        REncabezadoTrasladosPT2!Observaciones = UCase(REncabezadoTrasladosPT!Observaciones)
                        REncabezadoTrasladosPT2!Estado = UCase(REncabezadoTrasladosPT!Estado)
                    REncabezadoTrasladosPT2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Encabezado Traslados Producto Terminado"
                        Err.Clear
                    End If
                REncabezadoTrasladosPT.MoveNext
            Loop

            MsgBox "ya"

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RDetalleTrasladosPT = New ADODB.Recordset
            Call Abrir_Recordset(RDetalleTrasladosPT, "Select * From DetalleTrasladosProductoTerminado")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RDetalleTrasladosPT2 = New ADODB.Recordset
            Call Abrir_Recordset2(RDetalleTrasladosPT2, "Select * From DetalleTrasladosInventario")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RDetalleTrasladosPT.EOF
                        'BUSCA LA FECHA EN QUE ENTRO EL BULTO
                        Set RBuscaFechaEntrada = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaFechaEntrada, "Select BodegaATrasladar From EncabezadoTrasladosProductoTerminado Where Documento = " & RDetalleTrasladosPT!Documento)
                            
                            If RBuscaFechaEntrada.RecordCount > 0 Then
                                
                            Else
                                MsgBox "Busqueda Bodega A Trasladar No Existe", vbOKOnly + vbInformation, "Informacion"
                            End If
            
            
            
                    'AGREGA UN REGISTRO EN ORACLE
                    RDetalleTrasladosPT2.AddNew
                        RDetalleTrasladosPT2!Documento = RDetalleTrasladosPT!Documento
                        RDetalleTrasladosPT2!Tarima = RDetalleTrasladosPT!Tarima
                        RDetalleTrasladosPT2!FichaTecnica = UCase(RDetalleTrasladosPT!FichaTecnica)
                        RDetalleTrasladosPT2!CantidadSalida = RDetalleTrasladosPT!Cantidad
                        RDetalleTrasladosPT2!BodegaEntrada = RBuscaFechaEntrada(0) 'bodega a trasladar
                        RDetalleTrasladosPT2!DiferenciaReqCorMas = "0"
                        RDetalleTrasladosPT2!DiferenciaReqCor = "0"
                        RDetalleTrasladosPT2!CantidadDesperdicio = "0"
                        RDetalleTrasladosPT2!CantidadDesperdicioProveedor = "0"
                        RDetalleTrasladosPT2!CantidadReal = RDetalleTrasladosPT!Cantidad
                        RDetalleTrasladosPT2!Orden = ""
                        RDetalleTrasladosPT2!Fechaproduccion = RDetalleTrasladosPT!Fechaproduccion
                        RDetalleTrasladosPT2!LineaProduccion = RDetalleTrasladosPT!Linea
                                                
                    RDetalleTrasladosPT2.Update
                    If Err <> 0 And Err <> -2147467259 Then
                        MsgBox Err.Description & "Detalle Traslados Producto Terminado"
                        Err.Clear
                    ElseIf Err - 2147467259 Then
'                        MsgBox Err.Description & "Detalle Traslados Producto Terminado"
                        Err.Clear
                    End If
                RDetalleTrasladosPT.MoveNext
            Loop

            MsgBox "ya"
End Sub

Public Sub Nueve()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RAjustesMP = New ADODB.Recordset
            'Call Abrir_Recordset(RAjustesMP, "Select * From AjustesMateriaPrima")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RAjustesMP2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RAjustesMP2, "Select * From AjustesProductoTerminado")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RAjustesMP.EOF
                    
                    'BUSCA LA FECHA EN QUE ENTRO EL BULTO
            '            Set RBuscaFechaEntrada = New ADODB.Recordset
            '                Call Abrir_Recordset(RBuscaFechaEntrada, "Select E.FechaEntrada From EncabezadoEntradasMateriaPrima as E, DetalleEntradasMateriaPrima as D Where E.Documento = D.Documento And D.Codigo = '" & RAjustesMP!CodigoMateriaPrima & "' And D.NumeroIngreso = " & RAjustesMP!NumeroIngreso)
            '                If RBuscaFechaEntrada.RecordCount > 0 Then
            '                Else
            '                    MsgBox "Fecha De Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
            '                End If
            
            
                    'AGREGA UN REGISTRO EN ORACLE
           '         RAjustesMP2.AddNew
           '             RAjustesMP2!FechaOperacion = RAjustesMP!FechaOperacion
           '             RAjustesMP2!fecha = RAjustesMP!fecha
           '             RAjustesMP2!Documento = UCase(RAjustesMP!Documento)
           '             RAjustesMP2!Efecto = UCase(RAjustesMP!Efecto)
           '             RAjustesMP2!Fechaproduccion = RBuscaFechaEntrada(0)
           '             RAjustesMP2!Linea = "77"
           '             RAjustesMP2!FichaTecnica = UCase(RAjustesMP!CodigoMateriaPrima)
           '             RAjustesMP2!Tarima = UCase(RAjustesMP!NumeroIngreso)
           '             RAjustesMP2!Cantidad = UCase(RAjustesMP!Cantidad)
           '             RAjustesMP2!Observaciones = UCase(RAjustesMP!Observaciones)
           '             RAjustesMP2!Usuario = UCase(RAjustesMP!Usuario)
           '         RAjustesMP2.Update
                    
           '         If Err <> 0 Then
           '             MsgBox Err.Description & "Ajustes Materia Prima"
           '             Err.Clear
           '         End If
           '
           '     RAjustesMP.MoveNext
            'Loop

            'MsgBox "ya"
'______________________________________________________________________________________________________________________
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            'Set RAjustesPT = New ADODB.Recordset
            'Call Abrir_Recordset(RAjustesPT, "Select * From AjustesProductoTerminado")
            'ABRIMOS EL RECORDSET DE ORACLE
            'Set RAjustesPT2 = New ADODB.Recordset
            'Call Abrir_Recordset2(RAjustesPT2, "Select * From AjustesProductoTerminado")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            'Do Until RAjustesPT.EOF
                    'AGREGA UN REGISTRO EN ORACLE
            '        RAjustesPT2.AddNew
            '            RAjustesPT2(0) = RAjustesPT(0)
            '            RAjustesPT2(1) = RAjustesPT(1)
            '            RAjustesPT2(2) = UCase(RAjustesPT(2))
            '            RAjustesPT2(3) = UCase(RAjustesPT(3))
            '            RAjustesPT2(4) = RAjustesPT(4)
            '            RAjustesPT2(5) = UCase(RAjustesPT(5))
            '            RAjustesPT2(6) = UCase(RAjustesPT(6))
            '            RAjustesPT2(7) = RAjustesPT(7)
            '            RAjustesPT2(8) = RAjustesPT(8)
            '            RAjustesPT2(9) = UCase(RAjustesPT(9))
            '            RAjustesPT2(10) = UCase(RAjustesPT(10))
            '        RAjustesPT2.Update
            '
            '        If Err <> 0 Then
            '            MsgBox Err.Description & "Ajustes Producto Terminado"
            '            Err.Clear
            '        End If
            '
            '    RAjustesPT.MoveNext
            'Loop

End Sub


Public Sub diez()
On Error Resume Next
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RBatch = New ADODB.Recordset
            Call Abrir_Recordset(RBatch, "Select * From Batch where Month(Fec_Rut) = 04 And Year(fec_rut) = 2005")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RBatch2 = New ADODB.Recordset
            Call Abrir_Recordset2(RBatch2, "Select * From Batch")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RBatch.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RBatch2.AddNew
                        RBatch2(0) = RBatch(0)
                        RBatch2(1) = UCase(RBatch(1))
                        RBatch2(2) = RBatch(2)
                        RBatch2(3) = UCase(RBatch(3))
                        RBatch2(4) = UCase(RBatch(4))
                        RBatch2(5) = RBatch(5)
                        RBatch2(6) = UCase(RBatch(6))
                        RBatch2(7) = RBatch(7)
                    RBatch2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Batch"
                        Err.Clear
                    End If
                    
                RBatch.MoveNext
            Loop
            
'______________________________________________________________________________________________________________________
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RBatchDatos = New ADODB.Recordset
            Call Abrir_Recordset(RBatchDatos, "Select * From BatchDatos Where Batch > 8000 And Batch <= 9000")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RBatchDatos2 = New ADODB.Recordset
            Call Abrir_Recordset2(RBatchDatos2, "Select * From BatchDatos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RBatchDatos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RBatchDatos2.AddNew
                        RBatchDatos2(0) = UCase(RBatchDatos(0))
                        RBatchDatos2(1) = RBatchDatos(1)
                        RBatchDatos2(2) = RBatchDatos(2)
                        RBatchDatos2(3) = RBatchDatos(3)
                        RBatchDatos2(4) = RBatchDatos(4)
                        RBatchDatos2(5) = RBatchDatos(5)
                        RBatchDatos2(6) = RBatchDatos(6)
                        RBatchDatos2(7) = RBatchDatos(7)
                        RBatchDatos2(8) = RBatchDatos(8)
                        RBatchDatos2(9) = RBatchDatos(9)
                        RBatchDatos2(10) = RBatchDatos(10)
                        RBatchDatos2(11) = RBatchDatos(11)
                        RBatchDatos2(12) = UCase(RBatchDatos(12))
                    RBatchDatos2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Batch Datos"
                        Err.Clear
                    End If
                    
                RBatchDatos.MoveNext
            Loop
           
         
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RCapturaDesperdicioMP = New ADODB.Recordset
            Call Abrir_Recordset(RCapturaDesperdicioMP, "Select * From CapturaDesperdicioMateriaPrima where month(fecha) = 01 And year(fecha) = 2001")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RCapturaDesperdicioMP2 = New ADODB.Recordset
            Call Abrir_Recordset2(RCapturaDesperdicioMP2, "Select * From CapturaDesperdicio")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RCapturaDesperdicioMP.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RCapturaDesperdicioMP2.AddNew
                        RCapturaDesperdicioMP2(0) = RCapturaDesperdicioMP(0)
                        RCapturaDesperdicioMP2(1) = UCase(RCapturaDesperdicioMP(1))
                        RCapturaDesperdicioMP2(2) = UCase(RCapturaDesperdicioMP(2))
                        RCapturaDesperdicioMP2(3) = UCase(RCapturaDesperdicioMP(3))
                        RCapturaDesperdicioMP2(4) = UCase(RCapturaDesperdicioMP(4))
                        RCapturaDesperdicioMP2(5) = RCapturaDesperdicioMP(5)
                        RCapturaDesperdicioMP2(6) = RCapturaDesperdicioMP(6)
                        RCapturaDesperdicioMP2(7) = UCase(RCapturaDesperdicioMP(7))
                    RCapturaDesperdicioMP2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Captura Desperdicio MP"
                        Err.Clear
                    End If
                    
                RCapturaDesperdicioMP.MoveNext
            Loop
            
            MsgBox "ya"
'______________________________________________________________________________________________________________________
'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RCapturaRutinas = New ADODB.Recordset
            Call Abrir_Recordset(RCapturaRutinas, "Select * From CapturaRutinas where Fec_Rut >= #01/30/2002# And Fec_Rut <= #01/31/2002#")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RCapturaRutinas2 = New ADODB.Recordset
            Call Abrir_Recordset2(RCapturaRutinas2, "Select * From CapturaRutinas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RCapturaRutinas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RCapturaRutinas2.AddNew
                        RCapturaRutinas2(0) = RCapturaRutinas(0)
                        RCapturaRutinas2(1) = UCase(RCapturaRutinas(1))
                        RCapturaRutinas2(2) = UCase(RCapturaRutinas(2))
                        RCapturaRutinas2(3) = UCase(RCapturaRutinas(3))
                        RCapturaRutinas2(4) = RCapturaRutinas(4)
                        RCapturaRutinas2(5) = UCase(RCapturaRutinas(5))
                        RCapturaRutinas2(6) = RCapturaRutinas(6)
                        RCapturaRutinas2(7) = UCase(RCapturaRutinas(7))
                    RCapturaRutinas2.Update
                    
                    If Err <> 0 Then
'                       MsgBox Err.Description & "Captura Rutinas"
                        Err.Clear
                    End If
                    
                RCapturaRutinas.MoveNext
            Loop
            
          MsgBox "ya"
            

            
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set RCierreTarimas = New ADODB.Recordset
            Call Abrir_Recordset(RCierreTarimas, "Select * From CierreTarima") ' Where Month(Fecha) = 10 And Year(Fecha) = 2001")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set RCierreTarimas2 = New ADODB.Recordset
            Call Abrir_Recordset2(RCierreTarimas2, "Select * From CierreBulto")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until RCierreTarimas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    RCierreTarimas2.AddNew
                        RCierreTarimas2!fecha = RCierreTarimas!fecha
                        RCierreTarimas2!Turno = "1"
                        RCierreTarimas2!Linea = RCierreTarimas!Linea
                        RCierreTarimas2!BodegaSalida = RCierreTarimas!Bodega
                        RCierreTarimas2!Existencia = RCierreTarimas!Saldo
                        RCierreTarimas2!CantidadMas = 0
                        RCierreTarimas2!CantidadMenos = 0
                        RCierreTarimas2!ExistenciaNueva = RCierreTarimas!Saldo
                        RCierreTarimas2!ContadorInicial = 0
                        RCierreTarimas2!ContadorFinal = RCierreTarimas!Descargar
                        RCierreTarimas2!CantidadProcesada = RCierreTarimas!Descargar
                        RCierreTarimas2!DesperdicioProceso = RCierreTarimas!Desperdicio
                        RCierreTarimas2!DesperdicioProveedor = 0
                        RCierreTarimas2!CantidadProcesadaReal = RCierreTarimas!Descargar
                        RCierreTarimas2!Total = RCierreTarimas!Descargar
                        RCierreTarimas2!Usuario = UCase(RCierreTarimas!Usuario)
                        RCierreTarimas2!Fechaproduccion = RCierreTarimas!Fechaproduccion
                        RCierreTarimas2!LineaProduccion = RCierreTarimas!Linea
                        RCierreTarimas2!FichaTecnica = UCase(RCierreTarimas!FichaTecnica)
                        RCierreTarimas2!Tarima = RCierreTarimas!Tarima
                        RCierreTarimas2!Hora = "08:00"
                        RCierreTarimas2!Observaciones = RCierreTarimas!Observaciones
                    RCierreTarimas2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Cierre Tarimas"
                        Err.Clear
                    End If
                    
                RCierreTarimas.MoveNext
            Loop
            
            MsgBox "ya"
            
            
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REmpleadosCapturaCursos = New ADODB.Recordset
            Call Abrir_Recordset(REmpleadosCapturaCursos, "Select * From EmpleadosCapturaCursos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REmpleadosCapturaCursos2 = New ADODB.Recordset
            Call Abrir_Recordset2(REmpleadosCapturaCursos2, "Select * From EmpleadosCapturaCursos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REmpleadosCapturaCursos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REmpleadosCapturaCursos2.AddNew
                        REmpleadosCapturaCursos2(0) = UCase(REmpleadosCapturaCursos(0))
                        REmpleadosCapturaCursos2(1) = UCase(REmpleadosCapturaCursos(1))
                        REmpleadosCapturaCursos2(2) = REmpleadosCapturaCursos(2)
                        REmpleadosCapturaCursos2(3) = REmpleadosCapturaCursos(3)
                        REmpleadosCapturaCursos2(4) = UCase(REmpleadosCapturaCursos(4))
                    REmpleadosCapturaCursos2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Empleados Captura Cursos"
                        Err.Clear
                    End If
                    
                REmpleadosCapturaCursos.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REmpleadosCapturaAumentos = New ADODB.Recordset
            Call Abrir_Recordset(REmpleadosCapturaAumentos, "Select * From EmpleadosCapturaAumentos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REmpleadosCapturaAumentos2 = New ADODB.Recordset
            Call Abrir_Recordset2(REmpleadosCapturaAumentos2, "Select * From EmpleadosCapturaAumentos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REmpleadosCapturaAumentos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REmpleadosCapturaAumentos2.AddNew
                        REmpleadosCapturaAumentos2(0) = UCase(REmpleadosCapturaAumentos(0))
                        REmpleadosCapturaAumentos2(1) = REmpleadosCapturaAumentos(1)
                        REmpleadosCapturaAumentos2(2) = REmpleadosCapturaAumentos(2)
                        REmpleadosCapturaAumentos2(3) = UCase(REmpleadosCapturaAumentos(3))
                    REmpleadosCapturaAumentos2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Empleados Captura Aumentos"
                        Err.Clear
                    End If
                    
                REmpleadosCapturaAumentos.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REmpleadosCapturaFaltas = New ADODB.Recordset
            Call Abrir_Recordset(REmpleadosCapturaFaltas, "Select * From EmpleadosCapturaFaltas")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REmpleadosCapturaFaltas2 = New ADODB.Recordset
            Call Abrir_Recordset2(REmpleadosCapturaFaltas2, "Select * From EmpleadosCapturaFaltas")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REmpleadosCapturaFaltas.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REmpleadosCapturaFaltas2.AddNew
                        REmpleadosCapturaFaltas2(0) = UCase(REmpleadosCapturaFaltas(0))
                        REmpleadosCapturaFaltas2(1) = UCase(REmpleadosCapturaFaltas(1))
                        REmpleadosCapturaFaltas2(2) = REmpleadosCapturaFaltas(2)
                        REmpleadosCapturaFaltas2(3) = UCase(REmpleadosCapturaFaltas(3))
                    REmpleadosCapturaFaltas2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Empleados Captura Faltas"
                        Err.Clear
                    End If
                    
                REmpleadosCapturaFaltas.MoveNext
            Loop
'______________________________________________________________________________________________________________________

'______________________________________________________________________________________________________________________
            'ABRIMOS EL RECORDSET DE ACCESS
            Set REmpleadosHijos = New ADODB.Recordset
            Call Abrir_Recordset(REmpleadosHijos, "Select * From EmpleadosHijos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set REmpleadosHijos2 = New ADODB.Recordset
            Call Abrir_Recordset2(REmpleadosHijos2, "Select * From EmpleadosHijos")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Do Until REmpleadosHijos.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    REmpleadosHijos2.AddNew
                        REmpleadosHijos2(0) = UCase(REmpleadosHijos(0))
                        REmpleadosHijos2(1) = UCase(REmpleadosHijos(1))
                        REmpleadosHijos2(2) = REmpleadosHijos(2)
                        REmpleadosHijos2(3) = UCase(REmpleadosHijos(3))
                        REmpleadosHijos2(4) = UCase(REmpleadosHijos(4))
                    REmpleadosHijos2.Update
                    
                    If Err <> 0 Then
                        MsgBox Err.Description & "Empleados Hijos"
                        Err.Clear
                    End If
                    
                REmpleadosHijos.MoveNext
            Loop
'______________________________________________________________________________________________________________________


'ABRIMOS EL RECORDSET DE ACCESS
            Set RUsuarios = New ADODB.Recordset
            Call Abrir_Recordset(RUsuarios, "Select * From ReporteControlDeDespachos")
            'ABRIMOS EL RECORDSET DE ORACLE
            Set Rusuarios2 = New ADODB.Recordset
            Call Abrir_Recordset2(Rusuarios2, "Select * From ProductoEnTransito")
            'HACER HASTA QUE SEA FIN DE ARCHIVO EL DE ACCESSS
            Cont = 1
            Do Until RUsuarios.EOF
                    'AGREGA UN REGISTRO EN ORACLE
                    Rusuarios2.AddNew
                        Rusuarios2(0) = RUsuarios(0)
                        Rusuarios2(1) = RUsuarios(1)
                        Rusuarios2(2) = RUsuarios(2)
                        Rusuarios2(3) = UCase(RUsuarios(3))
                        Rusuarios2(4) = RUsuarios(4)
                        Rusuarios2(5) = UCase(RUsuarios(5))
                        Rusuarios2(6) = RUsuarios(6)
                        Rusuarios2(7) = RUsuarios(7)
                        Rusuarios2(8) = RUsuarios(8)
                        Rusuarios2(9) = RUsuarios(9)
                        Rusuarios2(10) = RUsuarios(10)
                        Rusuarios2(11) = Cont
                        
                    Rusuarios2.Update
                    If Err <> 0 Then
                        MsgBox Err.Description & "Usuarios"
                        Err.Clear
                    End If
                    
                    Cont = Cont + 1
                RUsuarios.MoveNext
            Loop


End Sub

Private Sub Form_DblClick()
        Set RCambiaDocumento = New ADODB.Recordset
            Call Abrir_Recordset(RCambiaDocumento, "Select * From EncabezadoEntradasMateriaPrima Order By Documento")
            Dim Cont As Integer
            Cont = 7102
            
                Do Until RCambiaDocumento.EOF
                            RCambiaDocumento!Documento = Cont
                            RCambiaDocumento.Update
                    RCambiaDocumento.MoveNext
                    Cont = Cont + 1
                Loop
End Sub

