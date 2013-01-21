VERSION 5.00
Begin VB.Form MigrarDatos 
   Caption         =   "MIgrarDatos"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt 
      Height          =   2295
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "MigrarDatos.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "migrar datos "
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "MigrarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cont As Integer
Dim RVariablesDescripcion As Recordset
Dim RVariablesDescripcion2 As Recordset

Dim RVariablesMedia As Recordset
Dim RVariablesMedia2 As Recordset

Dim RFichaTecnica As Recordset
Dim RFichaTecnica2 As Recordset

Dim RCorrelativosMateriaPrima As Recordset
Dim RCorrelativosMateriaPrima2 As Recordset

Dim RLineas As Recordset
Dim RLineas2 As Recordset

Dim RRutinas As Recordset
Dim RRutinas2 As Recordset

Dim RBatch As Recordset
Dim RBatch2 As Recordset

Dim RBatchDatos As Recordset
Dim RBatchDatos2 As Recordset

Dim RCapturaRutinas As Recordset
Dim RCapturaRutinas2 As Recordset

Dim RDefectos As Recordset
Dim RDefectos2 As Recordset

Dim RProveedores As Recordset
Dim RProveedores2 As Recordset

Dim RProduccion As Recordset
Dim RProduccion2 As Recordset

Dim RProduccionConDefectos As Recordset
Dim RProduccionConDefectos2 As Recordset

Dim RProduccionConMateriaPrima As Recordset
Dim RProduccionConMateriaPrima2 As Recordset

Dim RProduccionLiberada As Recordset
Dim RProduccionLiberada2 As Recordset

Dim RProduccionLiberadaConTarima As Recordset
Dim RProduccionLiberadaConTarima2 As Recordset

Dim RProduccionLiberadaConDefectos As Recordset
Dim RProduccionLiberadaConDefectos2 As Recordset

Dim RUsuarios As Recordset
Dim RUsuarios2 As Recordset

Dim RBultos As Recordset
Dim RBultos2 As Recordset

Dim RParos As Recordset
Dim RParos2 As Recordset

Dim RTraslados As Recordset
Dim RTraslados2 As Recordset

Dim RTrasladosDetalle As Recordset
Dim RTrasladosDetalle2 As Recordset

Dim REntradas As Recordset
Dim REntradas2 As Recordset

Dim REntradasDetalle As Recordset
Dim REntradasDetalle2 As Recordset

Dim REgresos As Recordset
Dim REgresos2 As Recordset

Dim REgresosDetalle As Recordset
Dim REgresosDetalle2 As Recordset


Private Sub Command1_Click()
On Error Resume Next

MousePointer = 11

    'VARIABLES DESCRIPCION
    Set RVariablesDescripcion = Db.OpenRecordset("Select * From VariablesDescripcion")
    Set RVariablesDescripcion2 = Db2.OpenRecordset("Select * From VariablesDescripcion")
    
    Do Until RVariablesDescripcion.EOF
            RVariablesDescripcion2.AddNew
                RVariablesDescripcion2(0) = RVariablesDescripcion(0)
                RVariablesDescripcion2(1) = RVariablesDescripcion(1)
            RVariablesDescripcion2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RVariablesDescripcion.MoveNext
    Loop
        Txt.Text = Txt.Text & "Variables Descripcion Listo" & vbCrLf
        
        
    'FICHA TECNICA
    Set RFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica")
    Set RFichaTecnica2 = Db2.OpenRecordset("Select * From FichaTecnica")
    
    Do Until RFichaTecnica.EOF
            RFichaTecnica2.AddNew
                RFichaTecnica2!Esp_Tec = RFichaTecnica!Esp_Tec
                RFichaTecnica2!Descrip = RFichaTecnica!Descrip
                RFichaTecnica2!Tipo = "1"
                RFichaTecnica2!Diametro = RFichaTecnica!Diametro
                RFichaTecnica2!Capacida = RFichaTecnica!Capacida
                RFichaTecnica2!Altura = RFichaTecnica!Altura
                RFichaTecnica2!Envases = RFichaTecnica!Envases
                RFichaTecnica2!Size = RFichaTecnica!Size
                RFichaTecnica2!Imp_Defe = RFichaTecnica!Imp_Defe
                RFichaTecnica2!Imp_Cali = RFichaTecnica!Imp_Cali
                RFichaTecnica2!Atributos = RFichaTecnica!Atributos
                RFichaTecnica2!Variables = RFichaTecnica!Variables
                RFichaTecnica2!Origen = RFichaTecnica!Origen
            RFichaTecnica2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RFichaTecnica.MoveNext
    Loop
        Txt.Text = Txt.Text & "Ficha Tecnica Listo" & vbCrLf
        
    'CORRELATIVOS MATERIA PRIMA
    Set RCorrelativosMateriaPrima = Db.OpenRecordset("Select * From CorrelativosMateriaPrima")
    Set RCorrelativosMateriaPrima2 = Db2.OpenRecordset("Select * From CorrelativosMateriaPrima")
    
    Do Until RCorrelativosMateriaPrima.EOF
            RCorrelativosMateriaPrima2.AddNew
                RCorrelativosMateriaPrima2!CodigoMateriaPrima = RCorrelativosMateriaPrima!CodigoMateriaPrima
                RCorrelativosMateriaPrima2!Descripcion = RCorrelativosMateriaPrima!Descripcion
                RCorrelativosMateriaPrima2!Correlativo = RCorrelativosMateriaPrima!Correlativo
                RCorrelativosMateriaPrima2!UnidadMedida = RCorrelativosMateriaPrima!UnidadMedida
                RCorrelativosMateriaPrima2!UnidadMedidaPeso = RCorrelativosMateriaPrima!UnidadMedidaPeso
                RCorrelativosMateriaPrima2!TipoDeMateriaPrima = RCorrelativosMateriaPrima!TipoDeMateriaPrima
                RCorrelativosMateriaPrima2!Espesor = RCorrelativosMateriaPrima!Espesor
                RCorrelativosMateriaPrima2!Minimo = RCorrelativosMateriaPrima!Minimo
                RCorrelativosMateriaPrima2!CuerposPorLamina = RCorrelativosMateriaPrima!CuerposPorLamina
            RCorrelativosMateriaPrima2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RCorrelativosMateriaPrima.MoveNext

    Loop
        Txt.Text = Txt.Text & "Correlativos Materia Prima Descripcion Listo" & vbCrLf
        

    'LINEAS
    Set RLineas = Db.OpenRecordset("Select * From Lineas")
    Set RLineas2 = Db2.OpenRecordset("Select * From Lineas")
    
    Do Until RLineas.EOF
            RLineas2.AddNew
                RLineas2!Linea = RLineas!Linea
                RLineas2!Descrip = RLineas!Descrip
                RLineas2!Esp_Tec = RLineas!Esp_Tec
                RLineas2!Activa = RLineas!Activa
                RLineas2!Tarima = RLineas!Tarima
                RLineas2!Velocidad = RLineas!Velocidad
                RLineas2!Orden = RLineas!Orden
                RLineas2!Grupo = RLineas!Grupo
            RLineas2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RLineas.MoveNext
    Loop
        Txt.Text = Txt.Text & "Lineas Listo" & vbCrLf
        

    'RUTINAS
    Set RRutinas = Db.OpenRecordset("Select * From Rutinas")
    Set RRutinas2 = Db2.OpenRecordset("Select * From Rutinas")
    
    Do Until RRutinas.EOF
            RRutinas2.AddNew
                RRutinas2!Rutina = RRutinas!Rutina
                RRutinas2!Descrip = RRutinas!Descrip
                RRutinas2!Cabezal = RRutinas!Cabezal
                RRutinas2!Imp_Rut = RRutinas!Imp_Rut
                RRutinas2!GeneraRutinaProceso = RRutinas!GeneraRutinaProceso
                RRutinas2!GeneraRutinaArranque = RRutinas!GeneraRutinaArranque
            RRutinas2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RRutinas.MoveNext
    Loop
        Txt.Text = Txt.Text & "Rutinas Listo" & vbCrLf
        
    'VARIABLES MEDIA
    Set RVariablesMedia = Db.OpenRecordset("Select * From VariablesMedia")
    Set RVariablesMedia2 = Db2.OpenRecordset("Select * From VariablesMedia")
    
    Do Until RVariablesMedia.EOF
            RVariablesMedia2.AddNew
                RVariablesMedia2(0) = RVariablesMedia(0)
                RVariablesMedia2(1) = RVariablesMedia(1)
                RVariablesMedia2(2) = RVariablesMedia(2)
                RVariablesMedia2(3) = RVariablesMedia(3)
                RVariablesMedia2(4) = RVariablesMedia(4)
                RVariablesMedia2(5) = RVariablesMedia(5)
            RVariablesMedia2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RVariablesMedia.MoveNext
    Loop
        Txt.Text = Txt.Text & "Variables Media Listo" & vbCrLf
    
    
    'BATCH
    Set RBatch = Db.OpenRecordset("Select * From Batch")
    Set RBatch2 = Db2.OpenRecordset("Select * From Batch")
    
    Do Until RBatch.EOF
            RBatch2.AddNew
                RBatch2!Batch = RBatch!Batch
                RBatch2!Linea = RBatch!Linea
                RBatch2!Fec_Rut = RBatch!Fec_Rut
                RBatch2!Hor_rut = RBatch!Hor_rut
                RBatch2!Esp_Tec = RBatch!Esp_Tec
                RBatch2!Cabezal = RBatch!Cabezal
                RBatch2!Rutina = RBatch!Rutina
                RBatch2!Valor = RBatch!Valor
            RBatch2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RBatch.MoveNext
    Loop
        Txt.Text = Txt.Text & "Batch Listo" & vbCrLf
    
    'BATCH DATOS
    Set RBatchDatos = Db.OpenRecordset("Select * From BatchDatos")
    Set RBatchDatos2 = Db2.OpenRecordset("Select * From BatchDatos")
    
    Do Until RBatchDatos.EOF
            RBatchDatos2.AddNew
                RBatchDatos2!Batch = RBatchDatos!Batch
                RBatchDatos2!Rutina = RBatchDatos!Rutina
                RBatchDatos2!Lim_Pro_In = RBatchDatos!Lim_Pro_In
                RBatchDatos2!Lim_Pro_Su = RBatchDatos!Lim_Pro_Su
                RBatchDatos2!Cv = RBatchDatos!Cv
                RBatchDatos2!LIM_Esp_IN = RBatchDatos!LIM_Esp_IN
                RBatchDatos2!Lim_Esp_Su = RBatchDatos!Lim_Esp_Su
                RBatchDatos2!CP = RBatchDatos!CP
                RBatchDatos2!Des_Std = RBatchDatos!Des_Std
                RBatchDatos2!Media = RBatchDatos!Media
                RBatchDatos2!Dat_Men = RBatchDatos!Dat_Men
                RBatchDatos2!Dat_May = RBatchDatos!Dat_May
                RBatchDatos2!Linea = RBatchDatos!Linea
            RBatchDatos2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RBatchDatos.MoveNext
    Loop
        Txt.Text = Txt.Text & "Batch Datos Listo" & vbCrLf
        
    'CAPTURA DE RUTINAS
    Set RCapturaRutinas = Db.OpenRecordset("Select * From CapturaRutinas")
    Set RCapturaRutinas2 = Db2.OpenRecordset("Select * From CapturaRutinas")
    
    Do Until RCapturaRutinas.EOF
            RCapturaRutinas2.AddNew
                RCapturaRutinas2!Linea = RCapturaRutinas!Linea
                RCapturaRutinas2!Fec_Rut = RCapturaRutinas!Fec_Rut
                RCapturaRutinas2!Hor_rut = RCapturaRutinas!Hor_rut
                RCapturaRutinas2!Esp_Tec = RCapturaRutinas!Esp_Tec
                RCapturaRutinas2!Cabezal = RCapturaRutinas!Cabezal
                RCapturaRutinas2!Rutina = RCapturaRutinas!Rutina
                RCapturaRutinas2!Valor = RCapturaRutinas!Valor
                RCapturaRutinas2!Catalogo = RCapturaRutinas!Catalogo
            RCapturaRutinas2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RCapturaRutinas.MoveNext
    Loop
        Txt.Text = Txt.Text & "Captura DE Rutinas Listo" & vbCrLf
        
    'DEFECTOS
    Set RDefectos = Db.OpenRecordset("Select * From Defectos")
    Set RDefectos2 = Db2.OpenRecordset("Select * From Defectos")
    
    Do Until RDefectos.EOF
            RDefectos2.AddNew
                RDefectos2!defecto = RDefectos!defecto
                RDefectos2!Descrip = RDefectos!Descrip
                RDefectos2!Tipo = RDefectos!Tipo
            RDefectos2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RDefectos.MoveNext
    Loop
        Txt.Text = Txt.Text & "Defectos Listo" & vbCrLf

    'PROVEEDORES
    Set RProveedores = Db.OpenRecordset("Select * From Proveedores")
    Set RProveedores2 = Db2.OpenRecordset("Select * From Proveedores")
    
    Do Until RProveedores.EOF
            RProveedores2.AddNew
                RProveedores2!TipoDeProveedor = RProveedores!TipoDeProveedor
                RProveedores2!CodigoProveedor = RProveedores!CodigoProveedor
                RProveedores2!Proveedor = RProveedores!Proveedor
                RProveedores2!Direccion = RProveedores!Direccion
                RProveedores2!Telefono = RProveedores!Telefono
                RProveedores2!Fax = RProveedores!Fax
                RProveedores2!Nit = RProveedores!Nit
                RProveedores2!Encargado = RProveedores!Encargado
            RProveedores2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RProveedores.MoveNext
    Loop
        Txt.Text = Txt.Text & "Proveedores Listo" & vbCrLf

    'PRODUCCION
    Set RProduccion = Db.OpenRecordset("Select * From Produccion")
    Set RProduccion2 = Db2.OpenRecordset("Select * From Produccion")
    
    Do Until RProduccion.EOF
            RProduccion2.AddNew
                RProduccion2!fec_prd = RProduccion!fec_prd
                RProduccion2!Hor_prd = RProduccion!Hor_prd
                RProduccion2!Tarima = RProduccion!Tarima
                RProduccion2!Linea = RProduccion!Linea
                RProduccion2!Esp_Tec = RProduccion!Esp_Tec
                RProduccion2!Batch = RProduccion!Batch
                RProduccion2!Envases = RProduccion!Envases
                RProduccion2!Calidad = RProduccion!Calidad
                RProduccion2!Muestra = RProduccion!Muestra
                RProduccion2!Turno = RProduccion!Turno
                RProduccion2!Cod_Emp = RProduccion!Cod_Emp
                RProduccion2!NoMP9301 = RProduccion!NoMP9301
                RProduccion2!ColorMP9301 = RProduccion!ColorMP9301
                RProduccion2!Orden = RProduccion!Orden
            RProduccion2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RProduccion.MoveNext
    Loop
        Txt.Text = Txt.Text & "Prodcuccion Listo" & vbCrLf

    'PRODUCCION CON MATERIA PRIMA (fondo)
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!Fondo) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!Fondo
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoFondo
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop
        

    'PRODUCCION CON MATERIA PRIMA (Hojalata)
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!Platina) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!Platina
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoHojalata
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop
        


    'PRODUCCION CON MATERIA PRIMA (Alambre)
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
         If IsNull(RProduccionConMateriaPrima!AlambreCobre) Then
         Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!AlambreCobre
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoAlambre
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
         End If
        RProduccionConMateriaPrima.MoveNext
    Loop



    'PRODUCCION CON MATERIA PRIMA (barniz liquido
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
        If IsNull(RProduccionConMateriaPrima!BarnizLiquido) Then
        Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!BarnizLiquido
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoBarnizL
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
        End If
        RProduccionConMateriaPrima.MoveNext
    Loop



    'PRODUCCION CON MATERIA PRIMA (sello solve
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!SelloSolve) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!SelloSolve
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoSelloSolve
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop

        
        
        
    'PRODUCCION CON MATERIA PRIMA (barniz polvo
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!BarnizPolvo) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!BarnizPolvo
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoBarnizP
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop

    
    'PRODUCCION CON MATERIA PRIMA (Nylon stre
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!NylonStrech) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!NylonStrech
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoNylon
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop


    

    'PRODUCCION CON MATERIA PRIMA (BOBINA
    Set RProduccionConMateriaPrima = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConMateriaPrima2 = Db2.OpenRecordset("Select * From ProduccionConMateriaPrima")
    
    Do Until RProduccionConMateriaPrima.EOF
          If IsNull(RProduccionConMateriaPrima!CodigoBobina) Then
          Else
            RProduccionConMateriaPrima2.AddNew
                RProduccionConMateriaPrima2!fec_prd = RProduccionConMateriaPrima!fec_prd
                RProduccionConMateriaPrima2!Linea = RProduccionConMateriaPrima!Linea
                RProduccionConMateriaPrima2!Esp_Tec = RProduccionConMateriaPrima!Esp_Tec
                RProduccionConMateriaPrima2!Tarima = RProduccionConMateriaPrima!Tarima
                RProduccionConMateriaPrima2!CodigoMateriaPrima = RProduccionConMateriaPrima!CodigoBobina
                RProduccionConMateriaPrima2!Bulto = RProduccionConMateriaPrima!CodigoIngresoBobina
            RProduccionConMateriaPrima2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                MsgBox Err.Number & " " & Err.Description
                
                Err.Clear
            End If
          End If
        RProduccionConMateriaPrima.MoveNext
    Loop



'PRODUCCION CON DEFECTOS (DEFECTO1
    Set RProduccionConDefectos = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConDefectos2 = Db2.OpenRecordset("Select * From ProduccionConDefectos")
    
    Do Until RProduccionConDefectos.EOF
          If IsNull(RProduccionConDefectos!defecto1) Then
          Else
            RProduccionConDefectos2.AddNew
                RProduccionConDefectos2!fec_prd = RProduccionConDefectos!fec_prd
                RProduccionConDefectos2!Linea = RProduccionConDefectos!Linea
                RProduccionConDefectos2!Esp_Tec = RProduccionConDefectos!Esp_Tec
                RProduccionConDefectos2!Tarima = RProduccionConDefectos!Tarima
                RProduccionConDefectos2!defecto = RProduccionConDefectos!defecto1
                RProduccionConDefectos2!Cantidad = RProduccionConDefectos!Cantidad1
            RProduccionConDefectos2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                'MsgBox Err.Number & " " & Err.Description
                
            End If
         End If
        RProduccionConDefectos.MoveNext
    Loop



    'PRODUCCION CON DEFECTOS (DEFECTO2
    Set RProduccionConDefectos = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConDefectos2 = Db2.OpenRecordset("Select * From ProduccionConDefectos")
    
    Do Until RProduccionConDefectos.EOF
        If IsNull(RProduccionConDefectos!defecto2) Then
        Else
            RProduccionConDefectos2.AddNew
                RProduccionConDefectos2!fec_prd = RProduccionConDefectos!fec_prd
                RProduccionConDefectos2!Linea = RProduccionConDefectos!Linea
                RProduccionConDefectos2!Esp_Tec = RProduccionConDefectos!Esp_Tec
                RProduccionConDefectos2!Tarima = RProduccionConDefectos!Tarima
                RProduccionConDefectos2!defecto = RProduccionConDefectos!defecto2
                RProduccionConDefectos2!Cantidad = RProduccionConDefectos!Cantidad2
            RProduccionConDefectos2.Update
            If Err.Number <> 0 And Err.Number = 3314 Then
                'MsgBox Err.Number & " " & Err.Description
                
            End If
        End If
        RProduccionConDefectos.MoveNext
    Loop

        
    'PRODUCCION CON DEFECTOS (DEFECTO 3
    Set RProduccionConDefectos = Db.OpenRecordset("Select * From Produccion")
    Set RProduccionConDefectos2 = Db2.OpenRecordset("Select * From ProduccionConDefectos")
    
    Do Until RProduccionConDefectos.EOF
          If IsNull(RProduccionConDefectos!defecto3) Then
          Else
            RProduccionConDefectos2.AddNew
                RProduccionConDefectos2!fec_prd = RProduccionConDefectos!fec_prd
                RProduccionConDefectos2!Linea = RProduccionConDefectos!Linea
                RProduccionConDefectos2!Esp_Tec = RProduccionConDefectos!Esp_Tec
                RProduccionConDefectos2!Tarima = RProduccionConDefectos!Tarima
                RProduccionConDefectos2!defecto = RProduccionConDefectos!defecto3
                RProduccionConDefectos2!Cantidad = RProduccionConDefectos!Cantidad3
            RProduccionConDefectos2.Update
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
                
                
            End If
         End If
        RProduccionConDefectos.MoveNext
    Loop
        
    
    
    'PRODUCCION LIBERADA
    Set RProduccionLiberada = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberada2 = Db2.OpenRecordset("Select * From ProduccionLiberada")
    
    Do Until RProduccionLiberada.EOF
            RProduccionLiberada2.AddNew
                RProduccionLiberada2!fec_prd = RProduccionLiberada!fec_prd
                RProduccionLiberada2!Hor_prd = RProduccionLiberada!Hor_prd
                RProduccionLiberada2!Tarima = RProduccionLiberada!Tarima
                RProduccionLiberada2!Linea = RProduccionLiberada!Linea
                RProduccionLiberada2!Esp_Tec = RProduccionLiberada!Esp_Tec
                RProduccionLiberada2!Batch = RProduccionLiberada!Batch
                RProduccionLiberada2!Envases = RProduccionLiberada!Envases
                RProduccionLiberada2!Calidad = RProduccionLiberada!Calidad
                RProduccionLiberada2!Turno = "1"
                RProduccionLiberada2!Cod_Emp = RProduccionLiberada!Cod_Emp
            RProduccionLiberada2.Update
            If Err.Number <> 0 And Err.Number <> 3022 Then
                MsgBox Err.Number & " " & Err.Description
            ElseIf Err.Number = 3022 Then
                cont = cont + 1
             End If
        RProduccionLiberada.MoveNext
    Loop
        



    'PRODUCCION LIBERADA CON TARIMAS (rechazados e incompletas)
    Set RProduccionLiberadaConTarima = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConTarima2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConTarimas")
    
    Do Until RProduccionLiberadaConTarima.EOF
            RProduccionLiberadaConTarima2.AddNew
                RProduccionLiberadaConTarima2!fec_prd = RProduccionLiberadaConTarima!fec_prd
                RProduccionLiberadaConTarima2!Tarima = RProduccionLiberadaConTarima!Tarima
                RProduccionLiberadaConTarima2!Linea = RProduccionLiberadaConTarima!Linea
                RProduccionLiberadaConTarima2!Esp_Tec = RProduccionLiberadaConTarima!Esp_Tec
                
                RProduccionLiberadaConTarima2!Fec_PrdL = RProduccionLiberadaConTarima!FechaTarIncRec
                RProduccionLiberadaConTarima2!LineaL = RProduccionLiberadaConTarima!LineaIncRec
                RProduccionLiberadaConTarima2!Esp_TecL = RProduccionLiberadaConTarima!FichaTecnicaIncRec
                RProduccionLiberadaConTarima2!TarimaL = RProduccionLiberadaConTarima!TarimaIncRec
                RProduccionLiberadaConTarima2!CalidadL = RProduccionLiberadaConTarima!CalidadRI
                
                RProduccionLiberadaConTarima2!Revisados = (RProduccionLiberadaConTarima!Desperdicio + RProduccionLiberada!EnvasesLiberados)
                RProduccionLiberadaConTarima2!NoConforme = RProduccionLiberadaConTarima!Desperdicio
                RProduccionLiberadaConTarima2!Liberados = RProduccionLiberadaConTarima!EnvasesLiberados
                RProduccionLiberadaConTarima2!EnTarima = RProduccionLiberadaConTarima!EnvasesLiberados
                
            RProduccionLiberadaConTarima2.Update
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConTarima.MoveNext
    Loop



    'PRODUCCION LIBERADA CON TARIMAS (complemento)
    Set RProduccionLiberadaConTarima = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConTarima2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConTarimas")
    
    Do Until RProduccionLiberadaConTarima.EOF
            RProduccionLiberadaConTarima2.AddNew
                RProduccionLiberadaConTarima2!fec_prd = RProduccionLiberadaConTarima!fec_prd
                RProduccionLiberadaConTarima2!Tarima = RProduccionLiberadaConTarima!Tarima
                RProduccionLiberadaConTarima2!Linea = RProduccionLiberadaConTarima!Linea
                RProduccionLiberadaConTarima2!Esp_Tec = RProduccionLiberadaConTarima!Esp_Tec
                
                RProduccionLiberadaConTarima2!Fec_PrdL = RProduccionLiberadaConTarima!FechaTarcom
                RProduccionLiberadaConTarima2!LineaL = RProduccionLiberadaConTarima!LineaCom
                RProduccionLiberadaConTarima2!Esp_TecL = RProduccionLiberadaConTarima!FichaTecnicaCom
                RProduccionLiberadaConTarima2!TarimaL = RProduccionLiberadaConTarima!TarimaCom
                RProduccionLiberadaConTarima2!CalidadL = "C"
                
                RProduccionLiberadaConTarima2!Revisados = RProduccionLiberadaConTarima!EnvCom
                RProduccionLiberadaConTarima2!NoConforme = "0"
                RProduccionLiberadaConTarima2!Liberados = RProduccionLiberadaConTarima!EnvCom
                RProduccionLiberadaConTarima2!EnTarima = RProduccionLiberadaConTarima!EnvCom
                
            RProduccionLiberadaConTarima2.Update
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConTarima.MoveNext
    Loop
        




    'PRODUCCION LIBERADA CON DEFECTOS
    Set RProduccionLiberadaConDefectos = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConDefectos2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConDefectos")
    
    Do Until RProduccionLiberadaConDefectos.EOF
        If IsNull(RProduccionLiberadaConDefectos!defecto1) Then
        Else
            RProduccionLiberadaConDefectos2.AddNew
                RProduccionLiberadaConDefectos2!fec_prd = RProduccionLiberadaConDefectos!fec_prd
                RProduccionLiberadaConDefectos2!Tarima = RProduccionLiberadaConDefectos!Tarima
                RProduccionLiberadaConDefectos2!Linea = RProduccionLiberadaConDefectos!Linea
                RProduccionLiberadaConDefectos2!Esp_Tec = RProduccionLiberadaConDefectos!Esp_Tec
                
                RProduccionLiberadaConDefectos2!Fec_PrdL = RProduccionLiberadaConDefectos!FechaTarIncRec
                RProduccionLiberadaConDefectos2!LineaL = RProduccionLiberadaConDefectos!LineaIncRec
                RProduccionLiberadaConDefectos2!Esp_TecL = RProduccionLiberadaConDefectos!FichaTecnicaIncRec
                RProduccionLiberadaConDefectos2!TarimaL = RProduccionLiberadaConDefectos!TarimaIncRec
                                
                RProduccionLiberadaConDefectos2!defecto = RProduccionLiberadaConDefectos!defecto1
                RProduccionLiberadaConDefectos2!Cantidad = RProduccionLiberadaConDefectos!Cantidad1
                
            RProduccionLiberadaConDefectos2.Update
        End If
            
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConDefectos.MoveNext
    Loop
        


'PRODUCCION LIBERADA CON DEFECTOS
    Set RProduccionLiberadaConDefectos = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConDefectos2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConDefectos")
    
    Do Until RProduccionLiberadaConDefectos.EOF
        If IsNull(RProduccionLiberadaConDefectos!defecto2) Then
        Else
            RProduccionLiberadaConDefectos2.AddNew
                RProduccionLiberadaConDefectos2!fec_prd = RProduccionLiberadaConDefectos!fec_prd
                RProduccionLiberadaConDefectos2!Tarima = RProduccionLiberadaConDefectos!Tarima
                RProduccionLiberadaConDefectos2!Linea = RProduccionLiberadaConDefectos!Linea
                RProduccionLiberadaConDefectos2!Esp_Tec = RProduccionLiberadaConDefectos!Esp_Tec
                
                RProduccionLiberadaConDefectos2!Fec_PrdL = RProduccionLiberadaConDefectos!FechaTarIncRec
                RProduccionLiberadaConDefectos2!LineaL = RProduccionLiberadaConDefectos!LineaIncRec
                RProduccionLiberadaConDefectos2!Esp_TecL = RProduccionLiberadaConDefectos!FichaTecnicaIncRec
                RProduccionLiberadaConDefectos2!TarimaL = RProduccionLiberadaConDefectos!TarimaIncRec
                                
                RProduccionLiberadaConDefectos2!defecto = RProduccionLiberadaConDefectos!defecto2
                RProduccionLiberadaConDefectos2!Cantidad = RProduccionLiberadaConDefectos!Cantidad2
                
            RProduccionLiberadaConDefectos2.Update
        End If
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConDefectos.MoveNext
    Loop



'PRODUCCION LIBERADA CON DEFECTOS
    Set RProduccionLiberadaConDefectos = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConDefectos2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConDefectos")
    
    Do Until RProduccionLiberadaConDefectos.EOF
        If IsNull(RProduccionLiberadaConDefectos!defecto3) Then
        Else
            RProduccionLiberadaConDefectos2.AddNew
                RProduccionLiberadaConDefectos2!fec_prd = RProduccionLiberadaConDefectos!fec_prd
                RProduccionLiberadaConDefectos2!Tarima = RProduccionLiberadaConDefectos!Tarima
                RProduccionLiberadaConDefectos2!Linea = RProduccionLiberadaConDefectos!Linea
                RProduccionLiberadaConDefectos2!Esp_Tec = RProduccionLiberadaConDefectos!Esp_Tec
                
                RProduccionLiberadaConDefectos2!Fec_PrdL = RProduccionLiberadaConDefectos!FechaTarIncRec
                RProduccionLiberadaConDefectos2!LineaL = RProduccionLiberadaConDefectos!LineaIncRec
                RProduccionLiberadaConDefectos2!Esp_TecL = RProduccionLiberadaConDefectos!FichaTecnicaIncRec
                RProduccionLiberadaConDefectos2!TarimaL = RProduccionLiberadaConDefectos!TarimaIncRec
                                
                RProduccionLiberadaConDefectos2!defecto = RProduccionLiberadaConDefectos!defecto3
                RProduccionLiberadaConDefectos2!Cantidad = RProduccionLiberadaConDefectos!Cantidad3
                
            RProduccionLiberadaConDefectos2.Update
        End If
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConDefectos.MoveNext
    Loop

Err.Clear

'PRODUCCION LIBERADA CON DEFECTOS
    Set RProduccionLiberadaConDefectos = Db.OpenRecordset("Select * From ProduccionTotal")
    Set RProduccionLiberadaConDefectos2 = Db2.OpenRecordset("Select * From ProduccionLiberadaConDefectos")
    
    Do Until RProduccionLiberadaConDefectos.EOF
        If IsNull(RProduccionLiberadaConDefectos!defecto4) Then
        Else
            RProduccionLiberadaConDefectos2.AddNew
                RProduccionLiberadaConDefectos2!fec_prd = RProduccionLiberadaConDefectos!fec_prd
                RProduccionLiberadaConDefectos2!Tarima = RProduccionLiberadaConDefectos!Tarima
                RProduccionLiberadaConDefectos2!Linea = RProduccionLiberadaConDefectos!Linea
                RProduccionLiberadaConDefectos2!Esp_Tec = RProduccionLiberadaConDefectos!Esp_Tec
                
                RProduccionLiberadaConDefectos2!Fec_PrdL = RProduccionLiberadaConDefectos!FechaTarIncRec
                RProduccionLiberadaConDefectos2!LineaL = RProduccionLiberadaConDefectos!LineaIncRec
                RProduccionLiberadaConDefectos2!Esp_TecL = RProduccionLiberadaConDefectos!FichaTecnicaIncRec
                RProduccionLiberadaConDefectos2!TarimaL = RProduccionLiberadaConDefectos!TarimaIncRec
                                
                RProduccionLiberadaConDefectos2!defecto = RProduccionLiberadaConDefectos!defecto4
                RProduccionLiberadaConDefectos2!Cantidad = RProduccionLiberadaConDefectos!Cantidad4
                
            RProduccionLiberadaConDefectos2.Update
        End If
            If Err.Number <> 0 Then
                'MsgBox Err.Number & " " & Err.Description
            End If
        RProduccionLiberadaConDefectos.MoveNext
    Loop

Err.Clear

    'USUARIOS
    Set RUsuarios = Db.OpenRecordset("Select * From Usuarios")
    Set RUsuarios2 = Db2.OpenRecordset("Select * From Usuarios")
    
    Do Until RUsuarios.EOF
            RUsuarios2.AddNew
                RUsuarios2(0) = RUsuarios(0)
                RUsuarios2(1) = RUsuarios(1)
                RUsuarios2(2) = RUsuarios(2)
            RUsuarios2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
                
            End If
        RUsuarios.MoveNext
    Loop
    '    MsgBox "USUARIOS Media Listo"
    
Err.Clear

    'BULTOS PROCESADOS
    Set RBultos = Db.OpenRecordset("Select * From NumerosIngresosProcesados")
    Set RBultos2 = Db2.OpenRecordset("Select * From NumerosIngresosProcesados")
    
    Do Until RBultos.EOF
            RBultos2.AddNew
                RBultos2(0) = RBultos(0)
                RBultos2(1) = RBultos(1)
                RBultos2(2) = RBultos(2)
                RBultos2(3) = RBultos(3)
                RBultos2(4) = RBultos(4)
                RBultos2(5) = RBultos(5)
                RBultos2(6) = RBultos(6)
                RBultos2(7) = RBultos(7)
                RBultos2(8) = RBultos(8)
                RBultos2(9) = RBultos(9)
                RBultos2(10) = RBultos(10)
                RBultos2(11) = RBultos(11)
                RBultos2(12) = RBultos(12)
                RBultos2(13) = RBultos(13)
                RBultos2(14) = RBultos(14)
                RBultos2(15) = RBultos(15)
                RBultos2(16) = RBultos(16)
                RBultos2(17) = RBultos(17)
                RBultos2(18) = RBultos(18)
                
            RBultos2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
'
            End If
        RBultos.MoveNext
    Loop
        
        
        
    
    'PAROS
    Set RParos = Db.OpenRecordset("Select * From Paros")
    Set RParos2 = Db2.OpenRecordset("Select * From Paros")
    
    Do Until RParos.EOF
            RParos2.AddNew
                RParos2(0) = RParos(0)
                RParos2(1) = RParos(1)
                RParos2(2) = RParos(2)
                RParos2(3) = RParos(3)
            RParos2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RParos.MoveNext
    Loop
        'MsgBox "Paros Listo"
    
    Err.Clear
    
    'ENCABEZDO DE TRASLADOS
    Set RTraslados = Db.OpenRecordset("Select * From EncabezadoTrasladosMateriaPrimaP")
    Set RTraslados2 = Db2.OpenRecordset("Select * From EncabezadoTrasladosMateriaPrimaP")
    
    Do Until RTraslados.EOF
            RTraslados2.AddNew
                RTraslados2(0) = RTraslados(0)
                RTraslados2(1) = RTraslados(1)
                RTraslados2(2) = RTraslados(2)
                RTraslados2(3) = RTraslados(3)
                RTraslados2(4) = RTraslados(4)
                RTraslados2(5) = RTraslados(5)
            RTraslados2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RTraslados.MoveNext
    Loop
        
    
    
    
    'DETALLE DE TRASLADOS
    
    Set RTrasladosDetalle = Db.OpenRecordset("Select * From DetalleTrasladosMateriaPrimaP")
    Set RTrasladosDetalle2 = Db2.OpenRecordset("Select * From DetalleTrasladosMateriaPrimaP")
    
    Do Until RTrasladosDetalle.EOF
            RTrasladosDetalle2.AddNew
                RTrasladosDetalle2(0) = RTrasladosDetalle(0)
                RTrasladosDetalle2(1) = RTrasladosDetalle(1)
                RTrasladosDetalle2(2) = RTrasladosDetalle(2)
                RTrasladosDetalle2(3) = RTrasladosDetalle(3)
                RTrasladosDetalle2(4) = RTrasladosDetalle(4)
                RTrasladosDetalle2(5) = RTrasladosDetalle(5)
                RTrasladosDetalle2(6) = RTrasladosDetalle(6)
                RTrasladosDetalle2(7) = RTrasladosDetalle(7)
                RTrasladosDetalle2(8) = RTrasladosDetalle(8)
                RTrasladosDetalle2(9) = RTrasladosDetalle(9)
                RTrasladosDetalle2(10) = RTrasladosDetalle(10)
                
            RTrasladosDetalle2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        RTrasladosDetalle.MoveNext
    Loop
        
        
        
    'ENCABEZADO DE ENTRADAS
    Set REntradas = Db.OpenRecordset("Select * From EncabezadoEntradasMateriaPrima")
    Set REntradas2 = Db2.OpenRecordset("Select * From EncabezadoEntradasMateriaPrima")
    
    Do Until REntradas.EOF
            REntradas2.AddNew
                REntradas2(0) = REntradas(0)
                REntradas2(1) = REntradas(1)
                REntradas2(2) = REntradas(2)
                REntradas2(3) = REntradas(3)
                REntradas2(4) = REntradas(4)
                REntradas2(5) = REntradas(5)
                REntradas2(6) = REntradas(6)
                REntradas2(7) = REntradas(7)
                REntradas2(8) = REntradas(8)
                REntradas2(9) = REntradas(9)
                REntradas2(10) = REntradas(10)
                REntradas2(11) = REntradas(11)
                REntradas2(12) = REntradas(12)
                REntradas2(13) = REntradas(13)
                REntradas2(14) = REntradas(14)
            REntradas2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        REntradas.MoveNext
    Loop
        
    
    
    'ENTRADAS DETALLE
    Set REntradasDetalle = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima")
    Set REntradasDetalle2 = Db2.OpenRecordset("Select * From DetalleEntradasMateriaPrima")
    
    Do Until REntradasDetalle.EOF
            REntradasDetalle2.AddNew
                REntradasDetalle2!Documento = REntradasDetalle!Documento
                REntradasDetalle2!Codigo = REntradasDetalle!Codigo
                REntradasDetalle2!Cantidad = REntradasDetalle!Cantidad
                REntradasDetalle2!NumeroUnicoSerieBoleta = REntradasDetalle!NumeroUnicoSerieBoleta
                REntradasDetalle2!OrdenBoleta = REntradasDetalle!OrdenBoleta
                REntradasDetalle2!BultoBoleta = REntradasDetalle!BultoBoleta
                REntradasDetalle2!FechaBoleta = REntradasDetalle!FechaBoleta
                REntradasDetalle2!BobinaBoleta = REntradasDetalle!BobinaBoleta
                REntradasDetalle2!NumeroIngreso = REntradasDetalle!NumeroIngreso
                If REntradasDetalle!Calidad = "ACEPTADA" Then
                    REntradasDetalle2!Calidad = "A"
                ElseIf REntradasDetalle!Calidad = "NO CONFORME" Then
                    REntradasDetalle2!Calidad = "R"
                ElseIf REntradasDetalle!Calidad = "ADVERTENCIA" Then
                    REntradasDetalle2!Calidad = "P"
                End If
                
                REntradasDetalle2!Observaciones = REntradasDetalle!Observaciones
                REntradasDetalle2!BodegaDisponibilidad = REntradasDetalle!BodegaDisponibilidad
                REntradasDetalle2!CantidadTraslado = REntradasDetalle!CantidadTraslado
                REntradasDetalle2!SaldoDisponibilidad = REntradasDetalle!SaldoDisponibilidad
                REntradasDetalle2!CantidadSalida = REntradasDetalle!CantidadSalida
                REntradasDetalle2!Peso = REntradasDetalle!Peso
                REntradasDetalle2!PesoEntrada = REntradasDetalle!PesoEntrada
                REntradasDetalle2!Estado = "I"
                
            REntradasDetalle2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        REntradasDetalle.MoveNext
    Loop
        
        
        
        
    'ENCABEZADO DE EGRESOS
    Set REgresos = Db.OpenRecordset("Select * From EncabezadoEgresosMateriaPrima")
    Set REgresos2 = Db2.OpenRecordset("Select * From EncabezadoEgresosMateriaPrima")
    
    Do Until REgresos.EOF
            REgresos2.AddNew
                REgresos2(0) = REgresos(0)
                REgresos2(1) = REgresos(1)
                REgresos2(2) = REgresos(2)
                REgresos2(3) = REgresos(3)
                REgresos2(4) = REgresos(4)
                REgresos2(5) = REgresos(5)
                REgresos2(6) = REgresos(6)
                REgresos2(7) = REgresos(7)
                REgresos2(8) = REgresos(8)
                REgresos2(9) = REgresos(9)
                REgresos2(10) = REgresos(10)
                REgresos2(11) = REgresos(11)
                REgresos2(12) = REgresos(12)
                REgresos2(13) = REgresos(13)
                REgresos2(14) = REgresos(14)
                REgresos2(15) = REgresos(15)
            REgresos2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        REgresos.MoveNext
    Loop
        
    
    
    'DETALLE DE EGRESOS
    Set REgresosDetalle = Db.OpenRecordset("Select * From DetalleEgresosMateriaPrima")
    Set REgresosDetalle2 = Db2.OpenRecordset("Select * From DetalleEgresosMateriaPrima")
    
    Do Until REgresosDetalle.EOF
            REgresosDetalle2.AddNew
                REgresosDetalle2(0) = REgresosDetalle(0)
                REgresosDetalle2(1) = REgresosDetalle(1)
                REgresosDetalle2(2) = REgresosDetalle(2)
                REgresosDetalle2(3) = REgresosDetalle(3)
                REgresosDetalle2(4) = REgresosDetalle(4)
            REgresosDetalle2.Update
            If Err.Number <> 0 Then
                MsgBox Err.Number & " " & Err.Description
            End If
        REgresosDetalle.MoveNext
    Loop
        
    

MousePointer = 0
MsgBox "PROCESO TERMINADO"
End Sub

Private Sub TabStrip1_Change()

End Sub
