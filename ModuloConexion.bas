Attribute VB_Name = "ModuloConexion"
Option Explicit

'CONEXION PARA ORACLE
Global Conexion As ADODB.Connection
Global ConexionSeametal As ADODB.Connection
Global ConexionCpa As ADODB.Connection
'Global Conexion2 As ADODB.Connection

Global StrSql As String
Global StrSqlSeaMetal As String
Global StrSqlCpa As String

'VARIABLE PARA CONECTARME AL TIPo DE PROVEEDOR
Global GTipoProveedor As String
Global GConectionString As String
Global GConectionStringSeaMetal As String
Global GConectionStringCpa As String

Global GConeccion As String
Global GPassword As String
Global GOrigenDeDatos As String

Global GPlanta As String

'VARIABLES PARA REPORTE
Global GNombreReporte As String
Global GCriteriaReporte As String
Global GTituloReporte As String
Global GComentarioReporte As String

'VARIABLES PARA BUSCAR BASE DE DATOS EN ARCHIVOS DE TEXTO
Global GRutaDeReportes As String
Global GRutaDeArchivosDeTexto As String

Global GRutaSeametal As String
Global GRutaSeametalChiapas As String
Global GRutaSeametalSanLuisPotosi As String

Global GRutaCpa As String
Global GRutaCpaSanLuisPotosi As String
Global GRutaCpaChiapas As String

Global GRutaEpa As String
Global GRutaEpaSanLuisPotosi As String
Global GRutaEpaChiapas As String


Public Sub Abrir_Recordset(Recordset As ADODB.Recordset, StrSql As String)
On Error Resume Next
    Recordset.ActiveConnection = Conexion
    Recordset.LockType = adLockOptimistic
    Recordset.CursorLocation = adUseClient
    Recordset.CursorType = adOpenDynamic
    Recordset.Open StrSql

    If Err <> 0 Then
    End If
    
End Sub
 
Public Sub Abrir_RecordsetSeaMetal(Recordset As ADODB.Recordset, StrSqlSeaMetal As String)
On Error Resume Next
    Recordset.ActiveConnection = ConexionSeametal
    Recordset.LockType = adLockOptimistic
    Recordset.CursorLocation = adUseClient
    Recordset.CursorType = adOpenDynamic
    Recordset.Open StrSqlSeaMetal

    If Err <> 0 Then
    End If
    
End Sub
 
Public Sub Abrir_RecordsetCpa(Recordset As ADODB.Recordset, StrSqlCpa As String)
On Error Resume Next
    Recordset.ActiveConnection = ConexionCpa
    Recordset.LockType = adLockOptimistic
    Recordset.CursorLocation = adUseClient
    Recordset.CursorType = adOpenDynamic
    Recordset.Open StrSqlCpa

    If Err <> 0 Then
    End If
    
End Sub
 

'Public Sub Abrir_Recordset2(Recordset2 As ADODB.Recordset, StrSql As String) ''

'    On Error Resume Next
'    Recordset2.ActiveConnection = Conexion2
'    Recordset2.LockType = adLockOptimistic
'    Recordset2.CursorLocation = adUseClient
'    Recordset2.CursorType = adOpenDynamic
'    Recordset2.Open StrSql
'    If Err <> 0 Then
'    End If
    '
'End Sub
 
Public Sub Desconectar()
    Conexion.Close
    Set Conexion = Nothing
End Sub
