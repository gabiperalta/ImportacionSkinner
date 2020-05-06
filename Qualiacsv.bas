Attribute VB_Name = "Qualiacsv"
Option Explicit

Public Sub ImportarQualiaCSV()

Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim ssqlProv As String
Dim rsCMPProv As New Recordset
Dim rsp As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v As String, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
'Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim sArchivo As String
Dim rsCMP As New Recordset
Dim vHoja As String
Dim lRowH As Long
Dim vtotalDeLineas As Long
Dim vdir As Integer
Dim regMod As Long

'Variables propias de Qualia
Dim vfIdDatosAdicionalesQualia As Long
Dim vfIdPoliza As Long
Dim vfCodigodeRamo As Long
Dim vfNroTarjetaDeCredito As String
Dim vfFechaModificacion As Date
Dim vfMotivoBaja As String
Dim vfDescripcionDelProducto As String
Dim vfMontoAsegurado As String
Dim vfFormaDePago As String
'Dim vfModulo As Long
Dim vfDescripcionDelModulo As String
Dim vfNacionalidad As String
Dim vfTipoTelefono1 As String
Dim vfTipoTelefono2 As String
Dim vfTelefono2 As String
Dim vfSucursalCuenta As String
'Dim vfDigitoVerificador As long
Dim vfPrima As String
Dim vfPremio As String
Dim vfCalleDelBienAsegurado As String
Dim vfNroDelBienAsegurado As String
Dim vfDptoDelBienAsegurado As String
Dim vfPisoDelBienAsegurado As String
Dim vfProvinciaDelBienAsegurado As String
Dim vfCPDelBienAsegurado As String
Dim vfLocalidadDelBienAsegurado As String
Dim vfTecho As String
Dim vfVentanas As String
Dim vfAntiguedad As String
Dim vfTipoCerradura As String
Dim vfDependencias As String
Dim vfAdicionalSeguroTecnico As String
Dim vfAdicionalPalosDeGolf As String
Dim vfAdicionalNotebook As String
Dim vfAdicionalLCD As String
Dim vfAdicionalConsola As String
Dim vfAdicionalConsolaOp1 As String
Dim vfAdicionalConsolaOp2 As String
Dim vfBeneficiario1Nombre As String
Dim vfBeneficiario1Apellido As String
Dim vfBeneficiario1IdTipoDoc As String
Dim vfBeneficiario1NroDocumento As String
Dim vfBeneficiario1Porcentaje As String
Dim vfBeneficiario1Mail As String
Dim vfBeneficiario2Nombre As String
Dim vfBeneficiario2Apellido As String
Dim vfBeneficiario2IdTipoDoc As String
Dim vfBeneficiario2NroDocumento As String
Dim vfBeneficiario2Porcentaje As String
Dim vfBeneficiario2Mail As String
Dim vfBeneficiario3Nombre As String
Dim vfBeneficiario3Apellido As String
Dim vfBeneficiario3IdTipoDoc As String
Dim vfBeneficiario3NroDocumento As String
Dim vfBeneficiario3Porcentaje As String
Dim vfBeneficiario3Mail As String
Dim vfIdEnCliente As String
'Dim vfCanal As Long
Dim vfFechaDeAlta As Date
Dim vfFechaDeBaja As String ' valores vacios si no estan dados de baja
'Dim vfCondicionIVA As Long
'Dim vfActividad As Long
Dim vfCodigoDeProducto As String
'Dim vfDescripcionDelVehiculo As String
'Dim vfNumeroDeMotor As String
'Dim VfNumeroDeChasis As String
'Dim vfPatenteDelVehiculo As String
'Dim vfAñoDelVehiculo As Long
'Dim vfValorDelVehiculo As Long
'Dim vfValorDeAccesorios As Long
Dim vfLugarDeNacimiento As String ' eb la base de prueba esta como string, antes estaba como long
Dim vfCalle As String 'estaba como long se pasa a asitring
Dim vfNumero As String
Dim vfDepartamento As String ' estaba como long se pasa a string ( como viene en la base de prueba)
Dim vfPiso As String
'Dim vfCuentaDeDebito As Long
'Dim vfNumeroDelUltimoRecibo As String
'Dim vfDeuda As Long
'Dim vfMarcaCilindroGNC As Long
'Dim vfNumeroDelregulador As Long
'Dim vfNumeroDeCilindroGNC As Long
'Dim vfMarca2CilindroGNC As Long
'Dim vfNumeroDe2CilindroGNC As Long
Dim vfBeneficiario1TipoDeDocumento As String
Dim vfBeneficiario1NumeroDeDocumento As String
Dim vfBeneficiario2TipoDeDocumento As String
Dim vfBeneficiario2NumeroDeDocumento As String
Dim vfBeneficiario3TipoDeDocumento As String
Dim vfNumeroDeCertificado As Long
Dim vfBeneficiario3NumeroDeDocumento As String
Dim vfNroDeTarjetaDeCredito As Long
Dim vfNumeroDelBienAsegurado As String
Dim vfNombreProductor As String
'Dim vfBANCO As String
Dim Ll As Long
Dim vFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim vLinea As Long
Dim vPosicion As Long
Dim vCampo As String

Dim fechaAux As String
Dim diaAux As String
Dim mesAux As String
Dim anoAux As String
Dim FechaInicial As Date
' -----------------------------------------

On Error Resume Next

Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
Dim camposParaValidar(106)
Dim vCS As String
'==========Aqui setear el Caracter de separacion=========
    vCS = "|"
'=======================================================

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

FechaInicial = Now()

    Ll = 0
    ll100 = 0
    vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & "Qualia.txt"
    If Not fs.FileExists(vFile) Then Exit Sub
    Set tf = fs.OpenTextFile(vFile, ForReading, True)
    
'======='control de lectura del archivo de datos=======================
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Exit Sub
    End If
'=====inicio del control de corrida====================================
    Dim rsCorr As New Recordset
    cn1.Execute "TM_CargaPolizasLogDeSetCorridasxcia " & lIdCia & ", " & vgCORRIDA
    ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcia = " & lIdCia & " and Registrosleidos is null"
    rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    If rsCorr.EOF Then
        MsgBox "no se determino la corrida, se detendra el proceso"
        Exit Sub
    Else
        vgCORRIDA = rsCorr("corrida")
    End If
'=======seteo control de lote================================================================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    Dim vControlDeModificados As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
'=================================================================================================

vtotalDeLineas = 0
vgidCia = lIdCia

cn.Execute "DELETE  FROM bandejadeentrada.dbo.importarQualiaCSV"


If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

          vdir = 0

   Do Until tf.AtEndOfStream
        Ll = Ll + 1
        vLinea = Ll
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*") 'ï»¿
        sLine = Replace(sLine, "ï", "") 'ï»¿
        sLine = Replace(sLine, "»", "") 'ï»¿
        sLine = Replace(sLine, "¿", "") 'ï»¿


          
            '====maneja los lotes para corte de importacion========
            nroLinea = nroLinea + 1
            If nroLinea = LongDeLote + 1 Then
                vLote = vLote + 1
                vControlDeModificados = 0
                nroLinea = 1
            End If
            '======================================================
        vPosicion = 0
    
'==================================================================================================
'este campo no aparece en el diseño del registro enviado.
        'vCampo = "campana"
        'vPosicion = vPosicion + 1
            'vcampana = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)

'If vcampana = "BSF" Then
      ' vgidCampana = 934  'SantaFe
    'ElseIf vcampana = "BSJ" Then
        'vgidCampana = 935 'SanJuan
    'ElseIf vcampana = "BERSA" Then
        'vgidCampana = 936 'EntreRios
    'ElseIf vcampana = "BER" Then
        'vgidCampana = 936 'EntreRios
    'ElseIf vcampana = "BSC" Then
        'vgidCampana = 937 'SantaCruz
    'End If
    
    
   '==================================================================================================
        vCampo = "nropoliza"
        vPosicion = vPosicion + 1
            vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)

    '==================================================================================================
       ' vCampo = "vgNROSECUENCIAL"
        'vPosicion = vPosicion + 1
            'vgNROSECUENCIAL = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
       'este campo no aparece en el diseño de registro.
   '==================================================================================================
        vCampo = "numero de certificados"
        vPosicion = vPosicion + 1
            vfNumeroDeCertificado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    
    '==================================================================================================
        vCampo = "vgCodigoDeRamo"
        vPosicion = vPosicion + 1
            vgCodigoDeRamo = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
          ssql = "select idproductomultiasistencias from tm_productosmultiasistencias where idcampana= " & lIdCampana & " and idProductoEnCliente = '" & vgCodigoDeRamo & "'"
        rsp.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        
        If Not rsp.EOF Then
            vgIdProducto = rsp("idproductomultiasistencias")
        Else
            vgIdProducto = -99
        End If
        rsp.Close
        
    '==================================================================================================
        vCampo = "Descipcion del ramo"
        vPosicion = vPosicion + 1
            vgOcupacion = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
         
     '==================================================================================================
'        vCampo = "sucursal"
'        vPosicion = vPosicion + 1
'            vgAgencia = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
'        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
'        vCampo = "Salta Campo vendedor"
'        vPosicion = vPosicion + 1
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'
        '==================================================================================================
        'vCampo = "canal"
        'vPosicion = vPosicion + 1
            'vfCanal = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
  
      '==================================================================================================
        vCampo = "nombre del cliente"
        vPosicion = vPosicion + 1
            vgNombre = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
  
      '==================================================================================================
        vCampo = "apellido del cliente"
        vPosicion = vPosicion + 1
            vgApellido = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
        '==================================================================================================
        vCampo = "id cliente"
        vPosicion = vPosicion + 1
            vfIdEnCliente = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
        '==================================================================================================
        'vCampo = "tipo de paquete"
        'vPosicion = vPosicion + 1
            'vgTipodeServicio = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        'vCampo = "numero de paquete"
        'vPosicion = vPosicion + 1
            'vgIdTipoDePOliza = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "tipo de documento"
        vPosicion = vPosicion + 1
            vgTipodeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "numero de documento"
        vPosicion = vPosicion + 1
            vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "sexo"
        vPosicion = vPosicion + 1
            vgSexo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "cuit/cuil cliente"
        vPosicion = vPosicion + 1
            vgCuit = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "numero de tarjeta de credito"
        'vPosicion = vPosicion + 1
            'vfNroTarjetaDeCredito = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "CBU"
        'vPosicion = vPosicion + 1
            'vfcbu = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "sucursal cuenta"
        'vPosicion = vPosicion + 1
            'vfSucursalCuenta = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "fecha de alta"
        vPosicion = vPosicion + 1
          fechaAux = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            diaAux = Mid(fechaAux, 7, 2)
            mesAux = Mid(fechaAux, 5, 2)
            anoAux = Mid(fechaAux, 1, 4)
            
            
            vgFECHAVIGENCIA = diaAux & "/" & mesAux & "/" & anoAux
           
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "fecha de ultima modificacion"
        vPosicion = vPosicion + 1
        
            fechaAux = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            diaAux = Mid(fechaAux, 7, 2)
            mesAux = Mid(fechaAux, 5, 2)
            anoAux = Mid(fechaAux, 1, 4)
            
            
            vfFechaModificacion = diaAux & "/" & mesAux & "/" & anoAux
            
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "fecha de baja "
        vPosicion = vPosicion + 1
        
          fechaAux = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            diaAux = Mid(fechaAux, 7, 2)
            mesAux = Mid(fechaAux, 5, 2)
            anoAux = Mid(fechaAux, 1, 4)
            
            
            vfFechaDeBaja = diaAux & "/" & mesAux & "/" & anoAux
           
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "motivo de baja"
        vPosicion = vPosicion + 1
            vfMotivoBaja = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "condicion iva"
        'vPosicion = vPosicion + 1
            'vfCondicionIVA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "actividad"
        'vPosicion = vPosicion + 1
            'vfActividad = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "descripcion del producto"
        vPosicion = vPosicion + 1
            vfDescripcionDelProducto = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
'            If InStr(1, vfDescripcionDelProducto, "BSF") Then
'               vgidCampana = 934  'SantaFe
'            ElseIf InStr(1, vfDescripcionDelProducto, "BSJ") Then
'                vgidCampana = 935 'SanJuan
'            ElseIf InStr(1, vfDescripcionDelProducto, "BERSA") Then
'                vgidCampana = 936 'EntreRios
'            ElseIf InStr(1, vfDescripcionDelProducto, "BER") Then
'                vgidCampana = 936 'EntreRios
'            ElseIf InStr(1, vfDescripcionDelProducto, "BSC") Then
'                vgidCampana = 937 'SantaCruz
'            End If
            
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "codigo de producto"
        vPosicion = vPosicion + 1
            vfCodigoDeProducto = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "monto asegurado"
        vPosicion = vPosicion + 1
            vfMontoAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "forma de pago"
        vPosicion = vPosicion + 1
            vfFormaDePago = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "descripcion del vehiculo"
        'vPosicion = vPosicion + 1
           ' vfDescripcionDelVehiculo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "modulo"
       ' vPosicion = vPosicion + 1
           ' vfModulo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "descripcion del modulo"
       ' vPosicion = vPosicion + 1
           ' vfDescripcionDelModulo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
       ' sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "numero de motor"
       'vPosicion = vPosicion + 1
            'vfNumeroDeMotor = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "numero de chasis"
       ' vPosicion = vPosicion + 1
           ' VfNumeroDeChasis = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "patente del vehiculo"
        'vPosicion = vPosicion + 1
           ' vfPatenteDelVehiculo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "año del vehiculo"
       ' vPosicion = vPosicion + 1
           ' vfAñoDelVehiculo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
      '  vCampo = "valor del vehiculo"
       ' vPosicion = vPosicion + 1
           ' vfValorDelVehiculo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
       ' sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
  
      '==================================================================================================
       ' vCampo = "valor de accesorios"
       ' vPosicion = vPosicion + 1
        '    vfValorDeAccesorios = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
 
      '==================================================================================================
        vCampo = "lugar de nacimiento"
        vPosicion = vPosicion + 1
            vfLugarDeNacimiento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "nacionalidad"
        vPosicion = vPosicion + 1
            vfNacionalidad = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "pais"
        vPosicion = vPosicion + 1
            vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "codigo postal"
        vPosicion = vPosicion + 1
            vgCODIGOPOSTAL = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "calle"
        vPosicion = vPosicion + 1
            vfCalle = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "numero"
        vPosicion = vPosicion + 1
            vfNumero = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "departamento"
        vPosicion = vPosicion + 1
            vfDepartamento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "piso"
        vPosicion = vPosicion + 1
            vfPiso = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "localidad"
        vPosicion = vPosicion + 1
            vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
        
      '==================================================================================================
        vCampo = "provincia"
        vPosicion = vPosicion + 1
            vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            ssqlProv = "Select denominacion from dicprovincias where idprovinciaqualia=" & vgPROVINCIA
               rsCMPProv.Open ssqlProv, cn1, adOpenForwardOnly, adLockReadOnly
            
               If Not rsCMPProv.EOF Then
                   vgPROVINCIA = rsCMPProv("denominacion")

               End If
               rsCMPProv.Close
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "mail cliente"
        vPosicion = vPosicion + 1
            vgEmail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "tipo  telefono1"
        vPosicion = vPosicion + 1
            vfTipoTelefono1 = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "telefono1"
        vPosicion = vPosicion + 1
            vgTelefono = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "tipo telefono 2"
        vPosicion = vPosicion + 1
            vfTipoTelefono2 = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "telefono2"
        vPosicion = vPosicion + 1
            vfTelefono2 = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
       ' vCampo = "cuenta de debito"
       ' vPosicion = vPosicion + 1
            'vfCuentaDeDebito = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        'vCampo = "sucursal cuenta "
        'vPosicion = vPosicion + 1
            'vfSucursalCuenta = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        'vCampo = "digito verificador"
       ' vPosicion = vPosicion + 1
           ' vfDigitoVerificador = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        'vCampo = "numero del ultimo recibo"
       ' vPosicion = vPosicion + 1
          '  vfNumeroDelUltimoRecibo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
      ' sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "prima"
        vPosicion = vPosicion + 1
            vfPrima = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
      '==================================================================================================
        vCampo = "premio"
        vPosicion = vPosicion + 1
            vfPremio = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
'        vCampo = "deuda"
'        vPosicion = vPosicion + 1
'            vfDeuda = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "calle del bien asegurado"
        vPosicion = vPosicion + 1
            vfCalleDelBienAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "numero del bien asegurado"
        vPosicion = vPosicion + 1
            vfNumeroDelBienAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "departamento del bien asegurado"
        vPosicion = vPosicion + 1
            vfDepartamento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "piso del bien asegurado"
        vPosicion = vPosicion + 1
            vfPiso = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "provincia del bien asegurado"
        vPosicion = vPosicion + 1
            vfProvinciaDelBienAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            'If vfProvinciaDelBienAsegurado <> 0 Then
             ssqlProv = "Select denominacion from dicprovincias where idprovinciaqualia=" & "'" & vfProvinciaDelBienAsegurado & "'"
               rsCMPProv.Open ssqlProv, cn1, adOpenForwardOnly, adLockReadOnly
            
               If Not rsCMPProv.EOF Then
                   vfProvinciaDelBienAsegurado = rsCMPProv("denominacion")

               End If
               rsCMPProv.Close
'               End If
               'vfProvinciaDelBienAsegurado = vfProvinciaDelBienAsegurado
               
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "codigo postal del bien asegurado"
        vPosicion = vPosicion + 1
            vfCPDelBienAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "localidad del bien asegurado"
        vPosicion = vPosicion + 1
            vfLocalidadDelBienAsegurado = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "techo"
        vPosicion = vPosicion + 1
            vfTecho = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "ventanas"
        vPosicion = vPosicion + 1
            vfVentanas = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "antiguedad años"
        vPosicion = vPosicion + 1
            vfAntiguedad = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "tipo cerradura"
        vPosicion = vPosicion + 1
            vfTipoCerradura = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "dependencias"
        vPosicion = vPosicion + 1
            vfDependencias = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional seguro tecnico (S/N)"
        vPosicion = vPosicion + 1
            vfAdicionalConsola = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional palos de golf (S/N)"
        vPosicion = vPosicion + 1
            vfAdicionalPalosDeGolf = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional notebook"
        vPosicion = vPosicion + 1
            vfAdicionalNotebook = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional lcd"
        vPosicion = vPosicion + 1
            vfAdicionalLCD = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional consola"
        vPosicion = vPosicion + 1
            vfAdicionalConsola = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional consola opcion 1 (S/N)"
        vPosicion = vPosicion + 1
            vfAdicionalConsolaOp1 = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "adicional consola opcion 2 (S/N)"
        vPosicion = vPosicion + 1
            vfAdicionalConsolaOp2 = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
  
      '==================================================================================================
       ' vCampo = "marca cilindro GNC"
       ' vPosicion = vPosicion + 1
            'vfMarcaCilindroGNC = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)


      '==================================================================================================
       ' vCampo = "numero del regulador"
       ' vPosicion = vPosicion + 1
           ' vfNumeroDelregulador = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "marca cilindro GNC"
       ' vPosicion = vPosicion + 1
         '   vfMarcaCilindroGNC = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
       ' sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "numero de cilindro GNC"
        'vPosicion = vPosicion + 1
        '    vfNumeroDeCilindroGNC = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "marca 2º cilindro GNC"
       ' vPosicion = vPosicion + 1
           ' vfMarca2CilindroGNC = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
       ' vCampo = "numero de 2º cilindro GNC "
       ' vPosicion = vPosicion + 1
           ' vfNumeroDe2CilindroGNC = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 nombre"
        vPosicion = vPosicion + 1
            vfBeneficiario1Nombre = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 apellido"
        vPosicion = vPosicion + 1
            vfBeneficiario1Apellido = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 tipo de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario1TipoDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 numero de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario1NumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 porcentaje"
        vPosicion = vPosicion + 1
            vfBeneficiario1Porcentaje = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario1 mail"
        vPosicion = vPosicion + 1
            vfBeneficiario1Mail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 nombre "
        vPosicion = vPosicion + 1
            vfBeneficiario2Nombre = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 apellido"
        vPosicion = vPosicion + 1
            vfBeneficiario2Apellido = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 tipo de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario2TipoDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 numero de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario2NumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 porcentaje"
        vPosicion = vPosicion + 1
            vfBeneficiario2Porcentaje = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario2 mail"
        vPosicion = vPosicion + 1
            vfBeneficiario2Mail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario3 nombre"
        vPosicion = vPosicion + 1
            vfBeneficiario3Nombre = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario3 apellido"
        vPosicion = vPosicion + 1
            vfBeneficiario3Apellido = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario3 tipo de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario3TipoDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario3 numero de documento"
        vPosicion = vPosicion + 1
            vfBeneficiario3NumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "Beneficiario3 porcentaje"
        vPosicion = vPosicion + 1
            vfBeneficiario3Porcentaje = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "beneficiario3 mail"
        vPosicion = vPosicion + 1
            vfBeneficiario3Mail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================

        vCampo = "NombreProductor"
'        vPosicion = vPosicion + 1
'            vgBANCO = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
            vPosicion = vPosicion + 1
            vfNombreProductor = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            If InStr(1, vfNombreProductor, "NUEVO BANCO SANTA FE S.A.") Then
               vgidCampana = 934  'SantaFe
            ElseIf InStr(1, vfNombreProductor, "BANCO DE SAN JUAN S.A.") Then
                vgidCampana = 935 'SanJuan
            ElseIf InStr(1, vfNombreProductor, "NUEVO BANCO ENTRE RIOS S.A.") Then
                vgidCampana = 936 'EntreRios
            ElseIf InStr(1, vfNombreProductor, "BANCO SANTA CRUZ S.A.") Then
                vgidCampana = 937 'SantaCruz
            ElseIf InStr(1, vfNombreProductor, "AGENTE DIRECTO") Then
                vgidCampana = 977 'Agente Directo
                vfIdEnCliente = vgNumeroDeDocumento
            Else
                vgidCampana = 932 'SantaCruz
            End If
            
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
      '==================================================================================================
        vCampo = "fecha de vigencia "
        vPosicion = vPosicion + 1

          'fechaAux = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
          fechaAux = Mid(sLine, 1)
            diaAux = Mid(fechaAux, 7, 2)
            mesAux = Mid(fechaAux, 5, 2)
            anoAux = Mid(fechaAux, 1, 4)


            vgFECHAVIGENCIA = diaAux & "/" & mesAux & "/" & anoAux

            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vgAPELLIDOYNOMBRE = vgApellido & ", " & vgNombre
        '==================================================================================================
        vgDOMICILIO = vfCalle & " " & vfNumero & " " & vfPiso & " " & vfDepartamento
        
        
'==============  IMPORTANTE   ================================================================.
'  Aqui controlamos si el registro ya existe en la base de datos de produccion
'   Si no existe hacemos el insert
'   si existe deberiamos comparar ciertos campos de la base de produccion con  los enviados,
'       si coinciden y no hay renovacion no se importa ese registro. De este modo achicamos la
'       cantidad de registros a importar y las actualizaciones que producen.
'       No podemos olvidar que los registros que no se importan los podria tomar el programa
'       como registros para dar de baja.
'   Con los registros que no ingresan deberiamos generar una lista o identificar estos registros
'       para avisar al programa que no los ponga de baja.
'   Todo esto se resuelve haciendo el control aqui cuando se importa en el temporal.
'       Indicando, en un campo "RegistroRepetido" para no importarlos pero que pueda ser usados
'       para indicar que habria que cambiarle en produccion la corrida y la fecha de corrida
'   .
Dim noSubidosaTemporal As Integer
noSubidosaTemporal = 0
    Dim rscn1 As New Recordset
    Dim rscn2 As New Recordset
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                
                ssql = " Select *  from tm_DatosAdicionalesQualia where idpoliza = " & rscn1("idpoliza")
                rscn2.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                If rscn2("FechaModificacion") = vfFechaModificacion Then
                    vdif = 0  'setea la variale de control de repetido con modificacion en cero
                Else
                    vControlDeModificados = vControlDeModificados + 1
                    If vControlDeModificados > 200 Then  ' este control es para evitar sobrecargar la impootacion cuando qualia haga una modificacion masiva.
                        vdif = 0
                    Else
                        vdif = 2  'setea la variale de control de repetido con modificacion en dos
                    End If
                End If
'                If Trim(rscn1("IDCIA")) <> Trim(vgidcia) Then vdif = vdif + 1
'                If Trim(rscn1("NUMEROCOMPANIA")) <> Trim(vgNUMEROCOMPANIA) Then vdif = vdif + 1
'                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
''                If Trim(rscn1("NROSECUENCIAL")) <> Trim(vgNROSECUENCIAL) Then vdif = vdif + 1
'                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
''                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
'                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
'                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
'                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
'                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
''                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
''                If Trim(rscn1("FECHAALTAOMNIA")) <> Trim(vgFECHAALTAOMNIA) Then vdif = vdif + 1
''                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
'                If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then vdif = vdif + 1
'                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
'                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
'                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
'                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
'                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
'                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
''                If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
'                If Trim(rscn1("IDTIPODECOBERTURA")) <> Trim(vgIDTIPODECOBERTURA) Then vdif = vdif + 1
''                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
''                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
'                If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
'                If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then vdif = vdif + 1
'                If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then vdif = vdif + 1
'                If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then vdif = vdif + 1
''                If Trim(rscn1("CORRIDA")) <> Trim(vgCORRIDA) Then vdif = vdif + 1
''                If Trim(rscn1("FECHACORRIDA")) <> Trim(vgFECHACORRIDA) Then vdif = vdif + 1
'                If Trim(rscn1("IdCampana")) <> Trim(vgidCampana) Then vdif = vdif + 1
'                If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
'                If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
'                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
'                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
'                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
'                If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then vdif = vdif + 1
'                If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then vdif = vdif + 1
'                If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then vdif = vdif + 1
'                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
'                If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
'                If Trim(rscn1("Gama")) <> Trim(vgGama) Then vdif = vdif + 1
'                If Trim(rscn1("InformadoSinCobertura")) <> Trim(vgInformadoSinCobertura) Then vdif = vdif + 1
'                If Trim(rscn1("MontoCoverturaVidrios")) <> Trim(vgMontoCoverturaVidrios) Then vdif = vdif + 1
''                If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
'                'If Trim(rscn1("CodigoDeProceso")) <> Trim(vgCodigoDeProceso) Then vdif = vdif + 1 no va
'                If Trim(rscn1("IdTipodePoliza")) <> Trim(vgIdTipoDePOliza) Then vdif = vdif + 1
''                If Trim(rscn1("Referido")) <> Trim(vgReferido) Then vdif = vdif + 1 ' no va
''                If Trim(rscn1("Telefono3")) <> Trim(vgTelefono3) Then vdif = vdif + 1' no va
'                If Trim(rscn1("IdProducto")) <> Trim(vgIdProducto) Then vdif = vdif + 1
'                If Trim(rscn1("email")) <> Trim(vgEmail) Then vdif = vdif + 1
'                If Trim(rscn1("email2")) <> Trim(vgemail2) Then vdif = vdif + 1
'                If Trim(rscn1("pais")) <> Trim(vgPais) Then vdif = vdif + 1
''                If Trim(rscn1("FechaNac")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1' no va
''                If Trim(rscn1("Modificaciones")) <> Trim(vgModificaciones) Then vdif = vdif + 1 ' no va
'                If Trim(rscn1("Sexo")) <> Trim(vgSexo) Then vdif = vdif + 1
''                If Trim(rscn1("Agencia")) <> Trim(vgAgencia) Then vdif = vdif + 1
'                If Trim(rscn1("Codigoencliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
'                If Trim(rscn1("DocumentoReferente")) <> Trim(vgDocumentoReferente) Then vdif = vdif + 1
'                If Trim(rscn1("Ocupacion")) <> Trim(vgOcupacion) Then vdif = vdif + 1
'                If Trim(rscn1("Cargo")) <> Trim(vgCargo) Then vdif = vdif + 1
'                If Trim(rscn1("Nombre")) <> Trim(vgNombre) Then vdif = vdif + 1
'                If Trim(rscn1("Apellido")) <> Trim(vgApellido) Then vdif = vdif + 1
''                If Trim(rscn1("Cuit")) <> Trim(vgCuit) Then vdif = vdif + 1
''                If Trim(rscn1("CBU")) <> Trim(vgCBU) Then vdif = vdif + 1' no va
''                If Trim(rscn1("SucursalCuenta")) <> Trim(vgSucursalCuenta) Then vdif = vdif + 1' no va
''                If Trim(rscn1("FechaModificacion")) <> Trim(vgFechaModificacion) Then vdif = vdif + 1
''                If Trim(rscn1("CondicionIVA")) <> Trim(vgCondicionIVA) Then vdif = vdif + 1' no va
''                If Trim(rscn1("DescripcionDelMundo")) <> Trim(vgDescripcionDelMundo) Then vdif = vdif + 1' no va
''                If Trim(rscn1("PaisOrigen")) <> Trim(vgPaisOrigen) Then vdif = vdif + 1
''                If Trim(rscn1("NroUltimoRecibo")) <> Trim(vgNroUltimoRecibo) Then vdif = vdif + 1' no va
''                If Trim(rscn1("DescripcionDelModulo")) <> Trim(vgDescripcionDelModulo) Then vdif = vdif + 1' no va

                vgIDPOLIZA = rscn1("idpoliza")
                rscn2.Close
            Else
                noSubidosaTemporal = noSubidosaTemporal + 1
            End If

        rscn1.Close
'-=================================================================================================================
 
        
    If Not IsDate(vfFechaDeBaja) Then
        ssql = "Insert into bandejadeentrada.dbo.importarQualiaCSV ("
        ssql = ssql & "IDPOLIZA,"
        ssql = ssql & "IDCIA,"
        ssql = ssql & "NUMEROCOMPANIA,"
        ssql = ssql & "NROPOLIZA,"
        ssql = ssql & "NROSECUENCIAL,"
        ssql = ssql & "APELLIDOYNOMBRE,"
        ssql = ssql & "DOMICILIO,"
        ssql = ssql & "LOCALIDAD,"
        ssql = ssql & "PROVINCIA,"
        ssql = ssql & "CODIGOPOSTAL,"
        ssql = ssql & "FECHAVIGENCIA,"
'        ssql = ssql & "FECHAVENCIMIENTO                  ,"
        ssql = ssql & "FECHAALTAOMNIA,"
        'ssql = ssql & "IDAUTO                            ,"
        'ssql = ssql & "MARCADEVEHICULO                   ,"
        'ssql = ssql & "MODELO                            ,"
        'ssql = ssql & "COLOR                             ,"
        'ssql = ssql & "ANO                               ,"
        'ssql = ssql & "PATENTE                           ,"
        'ssql = ssql & "TIPODEVEHICULO                    ,"
        ssql = ssql & "TipodeServicio,"
        'ssql = ssql & "IDTIPODECOBERTURA                 ,"
        ssql = ssql & "COBERTURAVEHICULO,"
        ssql = ssql & "COBERTURAVIAJERO,"
        ssql = ssql & "TipodeOperacion,"
        ssql = ssql & "Operacion,"
        ssql = ssql & "CATEGORIA,"
        ssql = ssql & "ASISTENCIAXENFERMEDAD,"
        ssql = ssql & "CORRIDA,"
        ssql = ssql & "FECHACORRIDA,"
        ssql = ssql & "IdCampana,"
        ssql = ssql & "Conductor,"
        ssql = ssql & "CodigoDeProductor,"
        ssql = ssql & "CodigoDeServicioVip,"
        ssql = ssql & "TipodeDocumento,"
        ssql = ssql & "NumeroDeDocumento,"
        ssql = ssql & "TipodeHogar,"
        ssql = ssql & "IniciodeAnualidad,"
        ssql = ssql & "PolizaIniciaAnualidad,"
        ssql = ssql & "Telefono,"
        ssql = ssql & "NroMotor,"
        ssql = ssql & "Gama,"
        ssql = ssql & "IdLote,"
        ssql = ssql & "InformadoSinCobertura,"
        'ssql = ssql & "MontoCoverturaVidrios             ,"
        ssql = ssql & "COBERTURAHOGAR,"
        ssql = ssql & "CodigoDeProceso,"
        'ssql = ssql & "IdTipodePoliza                    ,"
        ssql = ssql & "Referido,"
        ssql = ssql & "Telefono2,"
        ssql = ssql & "Telefono3,"
        ssql = ssql & "IdProducto,"
        ssql = ssql & "email,"
        ssql = ssql & "email2,"
        ssql = ssql & "pais,"
        ssql = ssql & "FechaNac,"
        ssql = ssql & "Sexo,"
        ssql = ssql & "Agencia,"
        ssql = ssql & "Codigoencliente,"
        ssql = ssql & "DocumentoReferente,"
        ssql = ssql & "Ocupacion,"
        'ssql = ssql & "Cargo                             ,"
        ssql = ssql & "Nombre,"
        ssql = ssql & "Apellido,"
        ssql = ssql & "IdEnCliente,"
        ssql = ssql & "Cuit,"
        'ssql = ssql & "NroDeTarjetaDeCredito             ,"
        ssql = ssql & "CBU,"
        'ssql = ssql & "SucursalCuenta                    ,"
        ssql = ssql & "FechaModificacion,"
        ssql = ssql & "MotivoBaja,"
        'ssql = ssql & "CondicionIVA                      ,"
        ssql = ssql & "DescripcionDelProducto,"
        ssql = ssql & "MontoAsegurado,"
        ssql = ssql & "FormaDePago,"
        ssql = ssql & "DescripcionDelMundo,"
        ssql = ssql & "PaisOrigen,"
        ssql = ssql & "Nacionalidad,"
        ssql = ssql & "TipoTelefono1,"
        ssql = ssql & "TipoTelefono2,"
        ssql = ssql & "Prima,"
        ssql = ssql & "Premio,"
        ssql = ssql & "CalleDelBienAsegurado,"
        ssql = ssql & "NumeroDelBienAsegurado,"
        ssql = ssql & "DptoDelBienAsegurado,"
        ssql = ssql & "ProvinciaDelBienAsegurado,"
        ssql = ssql & "CPDelBienAsegurado,"
        ssql = ssql & "LocalidadDelBienAsegurado,"
        ssql = ssql & "Techo,"
        ssql = ssql & "Ventanas,"
        ssql = ssql & "Antiguedad,"
        ssql = ssql & "TipoCerradura,"
        ssql = ssql & "Dependencias,"
        ssql = ssql & "AdicionalSeguroTecnico,"
        ssql = ssql & "AdicionalPalosDeGolf,"
        ssql = ssql & "AdicionalNotebook,"
        ssql = ssql & "AdicionalLCD,"
        ssql = ssql & "AdicionalConsolaOp1,"
        ssql = ssql & "AdicionalConsolaOp2,"
        ssql = ssql & "Beneficiario1Nombre,"
        ssql = ssql & "Beneficiario1Apellido,"
        ssql = ssql & "Beneficiario1IdTipoDoc,"
        ssql = ssql & "Beneficiario1NroDocumento,"
        ssql = ssql & "Beneficiario1Porcentaje,"
        ssql = ssql & "Beneficiario1Mail,"
        ssql = ssql & "Beneficiario2Nombre,"
        ssql = ssql & "Beneficiario2Apellido,"
        ssql = ssql & "Beneficiario2IdTipoDoc,"
        ssql = ssql & "Beneficiario2NroDocumento,"
        ssql = ssql & "Beneficiario2Porcentaje,"
        ssql = ssql & "Beneficiario2Mail,"
        ssql = ssql & "Beneficiario3Nombre,"
        ssql = ssql & "Beneficiario3Apellido,"
        ssql = ssql & "Beneficiario3IdTipoDoc,"
        ssql = ssql & "Beneficiario3NroDocumento,"
        ssql = ssql & "Beneficiario3Porcentaje,"
        ssql = ssql & "Beneficiario3Mail,"
        ssql = ssql & "NroTarjetaDeCredito,"
        ssql = ssql & "DescripcionDelModulo,"
        ssql = ssql & "NroDelBienAsegurado,"
        ssql = ssql & "PisoDelBienAsegurado,"
        ssql = ssql & "AdicionalConsola,"
        ssql = ssql & "CodigoDeProducto,"
        ssql = ssql & "CodigodeRamo, modificaciones)"
        ssql = ssql & " values("
          ssql = ssql & Trim(vgIDPOLIZA) & ", " ' variable de formulario propio de qualia
        ssql = ssql & Trim(vgidCia) & ", '"
        ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
        ssql = ssql & Trim(vgNROPOLIZA) & "', '"
        ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgLOCALIDAD) & "','"
        ssql = ssql & Trim(vgPROVINCIA) & "','"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
        ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
'        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgFECHAALTAOMNIA) & "', '"
        'ssql = ssql & Trim(vgIDAUTO) & ", '"
       ' ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
       ' ssql = ssql & Trim(vgMODELO) & "','"
       ' ssql = ssql & Trim(vgCOLOR) & "','"
        'ssql = ssql & Trim(vgAno) & "','"
        'ssql = ssql & Trim(vgPATENTE) & "', "
        'ssql = ssql & Trim(vgTIPODEVEHICULO) & ", '"
        ssql = ssql & Trim(vgTipodeServicio) & "', '"
        'ssql = ssql & Trim(vgIDTIPODECOBERTURA) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "','"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgTipodeOperacion) & "', '"
        ssql = ssql & Trim(vgOperacion) & "', '"
        ssql = ssql & Trim(vgCATEGORIA) & "', '"
        ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', "
        ssql = ssql & Trim(vgCORRIDA) & ", '"
        ssql = ssql & Trim(vgFECHACORRIDA) & "', "
        ssql = ssql & Trim(vgidCampana) & ", '"
        ssql = ssql & Trim(vgConductor) & "', '"
        ssql = ssql & Trim(vgCodigoDeProductor) & "','"
        ssql = ssql & Trim(vgCodigoDeServicioVip) & "','"
        ssql = ssql & Trim(vgTipodeDocumento) & "','"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "','"
        ssql = ssql & Trim(vgTipodeHogar) & "','"
        ssql = ssql & Trim(vgIniciodeAnualidad) & "', '"
        ssql = ssql & Trim(vgPolizaIniciaAnualidad) & "', '"
        ssql = ssql & Trim(vgTelefono) & "', '"
        ssql = ssql & Trim(vgNroMotor) & "', '"
        ssql = ssql & Trim(vgGama) & "', "
        ssql = ssql & Trim(vLote) & ",'"
        ssql = ssql & Trim(vgInformadoSinCobertura) & "', '"
        'ssql = ssql & Trim(vgMontoCoverturaVidrios) & ", '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
        ssql = ssql & Trim(vgCodigoDeProceso) & "', '"
        'ssql = ssql & Trim(vgIdTipoDePOliza) & ", '"
        ssql = ssql & Trim(vgReferido) & "', '"
        ssql = ssql & Trim(vfTelefono2) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vgTelefono3) & "', "
        ssql = ssql & Trim(vgIdProducto) & ",'"
        ssql = ssql & Trim(vgEmail) & "', '"
        ssql = ssql & Trim(vgEmail2) & "','"
        ssql = ssql & Trim(vgPais) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "','"
        ssql = ssql & Trim(vgSexo) & "', '"
        ssql = ssql & Trim(vgAgencia) & "', '"
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
        ssql = ssql & Trim(vgDocumentoReferente) & "', '"
        ssql = ssql & Trim(vgOcupacion) & "', '"
        'ssql = ssql & Trim(vgCargo) & ",'"
        ssql = ssql & Trim(vgNombre) & "','"
        ssql = ssql & Trim(vgApellido) & "', '"
        ssql = ssql & Trim(vfIdEnCliente) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vgCuit) & "', '"
       ' ssql = ssql & Trim(vfNroDeTarjetaDeCredito) & ", '" ' variable de formulario de qualia
        ssql = ssql & Trim(vgCBU) & "', '"
        'ssql = ssql & Trim(vgSucursalCuenta) & ", '"
        ssql = ssql & Trim(vfFechaModificacion) & "','"
        ssql = ssql & Trim(vfMotivoBaja) & "', '" ' variable de formulario de qualia
        'ssql = ssql & Trim(vgCondicionIVA) & ", '"
        ssql = ssql & Trim(vfDescripcionDelProducto) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfMontoAsegurado) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfFormaDePago) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vgDescripcionDelMundo) & "', '"
        ssql = ssql & Trim(vgPaisOrigen) & "','"
        ssql = ssql & Trim(vfNacionalidad) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfTipoTelefono1) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfTipoTelefono2) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfPrima) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfPremio) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfCalleDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfNumeroDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfDptoDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfProvinciaDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfCPDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfLocalidadDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfTecho) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfVentanas) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAntiguedad) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfTipoCerradura) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfDependencias) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalSeguroTecnico) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalPalosDeGolf) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalNotebook) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalLCD) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalConsolaOp1) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalConsolaOp2) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1Nombre) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1Apellido) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1IdTipoDoc) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1NroDocumento) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1Porcentaje) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario1Mail) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2Nombre) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2Apellido) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2IdTipoDoc) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2NroDocumento) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2Porcentaje) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario2Mail) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3Nombre) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3Apellido) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3IdTipoDoc) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3NroDocumento) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3Porcentaje) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfBeneficiario3Mail) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfNroTarjetaDeCredito) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vgDescripcionDelModulo) & "','"
        ssql = ssql & Trim(vfNroDelBienAsegurado) & "', '" ' variable de formulario de qualia
        ssql = ssql & Trim(vfPisoDelBienAsegurado) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfAdicionalConsola) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vfCodigoDeProducto) & "','" ' variable de formulario de qualia
        ssql = ssql & Trim(vgCodigoDeRamo) & "'," ' variable de formulario de qualia
        ssql = ssql & Trim(vdif) & ")   "
                cn.Execute ssql
        End If
        

'========Control de errores=========================================================
                        If Err Then
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, "")
                            Err.Clear
                        
                        End If
'===========================================================================================
      
        vdir = 0
              
        If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        'Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
        
        ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
            ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcia=" & vgidCia & "and corrida =" & vgCORRIDA
            cn1.Execute ssql
            
            ll100 = 0
        End If
        DoEvents
              
              
              DoEvents
              
              
          Loop
    
'    'Else
'        MsgBox ("Los siguientes campos obligatorios de la hoja nro " & j & " no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
'        oExcel.Workbooks.Close
'        Set oExcel = Nothing
'
'        Exit Sub
'    End If
    


    
'================Control de Leidos===============================================
                            cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
                           ' listoParaProcesar
'=================================================================================


    
    ImportadordePolizas.txtprocesando.Text = " Procesando un total de " & Ll & " datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos===========================================
                            cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
'    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
'    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
'    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    
    cn1.Execute "update tm_campana set UltimaCorridaError='' , UltimaCorridaCantidadderegistros=0  where idcontacto= " & vgidCia
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    DoEvents
    
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    DoEvents
'    cn1.CommandTimeout = 300
'    cn1.Execute sSPImportacion
'
'    ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & vgidcampana
'    rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'    vRegistrosProcesados = vRegistrosProcesados + rsCMP("UltimaCorridaCantidadImportada")
'    If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
'        MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
'    End If
'    rsCMP.Close
    
     For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
'            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
'            lLote = vLote + 1 'para salir del FOR
'        Else
                ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
'        End If

        rsCMP.Close
    Next lLote

'============Finaliza Proceso========================================================
                            cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxcia " & lIdCia & ", " & vgCORRIDA
                            Procesado
'=====================================================================================
   
    
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF
    
    
    cn1.Execute "update tm_campana set  UltimaCorridaCantidadderegistros = " & vRegistrosProcesados & " where idcontacto=" & vgidCia

'oExcel.Workbooks.Close
'Set oExcel = Nothing

Exit Sub
errores:
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow & " Columna: " & sName
    End If



End Sub


