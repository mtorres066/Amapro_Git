Attribute VB_Name = "Parametros"
Option Explicit

'Global Db As Database
'Global Db2 As Database
Global DBSeaMetal As Database
Global DBCpa As Database

Global GUsuario As String
Global BasedeDatos As String

'VARIABLES PARA LA CAPTURA DE PRODUCCION INTERNA
'Y TAMBIEN LA CAPTURA DE DEFECTOS DE PRODUCCION
Public VPLinea As String
Public VPFicha As String
Public VPTarima As Long
Public VPFecha As Date
Public VPCalidad As String

'VARIABLES PARA LA CAPTURA DE PRODUCCION INTERNA
'Y TAMBIEN LA CAPTURA DE DEFECTOS DE PRODUCCION
Public VPLLinea As String
Public VPLFicha As String
Public VPLTarima As Long
Public VPLFecha As Date
Public VPLCalidad As String

'VARIABLES PARA LA CAPTURA DE DEFECTOS DE PRODUCCION LIBERADA
Public VPDLinea As String
Public VPDFicha As String
Public VPDTarima As Long
Public VPDFecha As Date

Public VPLDLinea As String
Public VPLDFicha As String
Public VPLDTarima As Long
Public VPLDFecha As Date

'CALIDAD Y PRODUCCION
Global GConfiguracionCalidad As Boolean
Global GProduccion As Boolean
Global GEspecificaciones As Boolean
Global GReportesCalidad As Boolean
Global GAjustesInventario As Boolean

'EFICIENCIA
Global GConfiguracionEficiencia As Boolean
Global GCapturaParos As Boolean
Global GReportesEficiencia As Boolean
Global GEditarEficiencia As Boolean
Global GBorrarEficiencia As Boolean


'GENERALES
Global GUsuarios As Boolean
Global GEditar As Boolean
Global GBorrar As Boolean

'ORDENDES DE PRODUCCION
Global GOrdenProduccion As Boolean
Global GInvVenRepEje As Boolean
Global GReportesOrdenes As Boolean

'INVENTARIO
Global GConfiguracionInventario As Boolean
Global GEntradas As Boolean
Global GTraslados As Boolean
Global GSalidas As Boolean
Global GCambiosUbicacion As Boolean
Global GLiberacionEntradas As Boolean
Global GLiberacionSalidas As Boolean
Global GLiberacionTraslados As Boolean
Global GCierreBulto As Boolean
Global GReportesInventario As Boolean
Global GGraficasInventario As Boolean
Global GCapturaTransito As Boolean
Global GConsultaTransito As Boolean
Global GPorConEntInv As Boolean
Global GReportesFormatos As Boolean
Global GCapturaDesperdicio As Boolean
Global GReclamosProveedor As Boolean
Global GInspeccion As Boolean
                    
                    
                    

'PEDIDOS
Global GPedidosClientes As Boolean
Global GPedidosProveedores As Boolean
Global GCierreClientes As Boolean
Global GCierreProveedores As Boolean
Global GEditarPedidos As Boolean
Global GBorrarPedidos As Boolean


'EMPLEADOS
Global GConfiguracionEmpleados As Boolean
Global GCapturaFaltas As Boolean
Global GCapturaCursos As Boolean
Global GCapturaAumentos As Boolean
Global GReportesEmpleados As Boolean



Global VSql As String


Public Function FormatSingle(Numero As Single) As String
Dim SNumero As String
    SNumero = Space(15) & Format(Numero, "#,###,##0.00")
    FormatSingle = Right(SNumero, 16)
End Function

Public Function FormatInteger5(Numero2 As Integer) As String
Dim SNumero2 As String
    SNumero2 = Space(5) & Format(Numero2, "#,###,##0")
    FormatInteger5 = Right(SNumero2, 6)
End Function

Public Function FormatString15(Caracter As String) As String
Dim Scaracter As String
    Scaracter = Caracter & Space(15)
    FormatString15 = Left(Scaracter, 15)
End Function

Public Function UltimoDiaMes(fecha As Date)
Dim VDiasMes As Integer

    If Month(fecha) = "1" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2003" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2004" Then
            VDiasMes = "29"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2005" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2006" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2007" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2008" Then
            VDiasMes = "29"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2009" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2010" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2011" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2012" Then
            VDiasMes = "29"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2013" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2014" Then
            VDiasMes = "28"
        ElseIf Month(fecha) = "2" And Year(fecha) = "2015" Then
            VDiasMes = "28"
            
            
        ElseIf Month(fecha) = "3" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "4" Then
            VDiasMes = "30"
        ElseIf Month(fecha) = "5" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "6" Then
            VDiasMes = "30"
        ElseIf Month(fecha) = "7" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "8" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "9" Then
            VDiasMes = "30"
        ElseIf Month(fecha) = "10" Then
            VDiasMes = "31"
        ElseIf Month(fecha) = "11" Then
            VDiasMes = "30"
        ElseIf Month(fecha) = "12" Then
            VDiasMes = "31"
        End If
        
        UltimoDiaMes = VDiasMes
End Function


