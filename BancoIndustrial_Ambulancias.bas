Attribute VB_Name = "BancoIndustrial_Ambulancias"
Public Sub ImportarBancoIndustrial_Ambulancias()
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim vCampo As String
Dim vPosicion As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
Dim vlineasTotales As Long
Dim vCalle As String
Dim vAltura As String
Dim vPiso As String
Dim vDepto As String
Dim vTorre As String
Dim vVigenciaVigente As Date
Dim vBarrio As String
Dim regMod As Long
Dim vFecha As Date
Dim vgDia As Integer
Dim vgMes As Integer
Dim vgAno As Integer
Dim fechaActual As Date
Dim vDiferenciaFecha As Long
    
fechaActual = Now
'----------------
Dim vCS As String
vCS = ";"
'----------------

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

TablaTemporal 'se crea tabla temporal

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)

'======='control de lectura del archivo de datos=======================
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Exit Sub
    End If
'=====inicio del control de corrida====================================
    Dim rsCorr As New Recordset
    cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lidCampana & ", " & vgCORRIDA
    ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcampana = " & lidCampana & " and Registrosleidos is null"
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
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
'=================================================================================================
    Ll = 1
    tf.SkipLine 'Saltea la linea de encabezados
    Do Until tf.AtEndOfStream
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
        
'=======Control de Lote===============================
        nroLinea = nroLinea + 1

        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'=====================================================
Blanquear
vFecha = "00:00:00"
'Sucursal
        vCampo = "Sucursal"
        vPosicion = 1
        vgAgencia = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Nombre
        vCampo = "Nombre"
        vPosicion = 1
        vgNombre = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Apellido
        vCampo = "Apellido"
        vPosicion = 1
        vgApellido = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Tipo Beneficio
'        vCampo = "Tipo Beneficio"
'        vPosicion = 1
'        vgCargo = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Tipo Documento
'        vCampo = "Tipo Documento"
'        vPosicion = 1
'        vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Nro Documento
        vCampo = "Nro Documento"
        vPosicion = 1
        vgNumeroDeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        vgNROPOLIZA = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'SEXO
        vCampo = "SEXO"
        vPosicion = 1
        vgSexo = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'FechaNacimiento
        vCampo = "FechaNacimiento"
        vPosicion = 1
        vgFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Cuit
        vCampo = "Cuit"
        vPosicion = 1
        vgCuit = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        vgNROPOLIZA = vgCuit
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'IVA
'        vCampo = "IVA"
'        vPosicion = 1
'        vgTipodeOperacion = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
'        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Calle
        vCampo = "Calle"
        vPosicion = 1
        vCalle = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Altura
        vCampo = "Altura"
        vPosicion = 1
        vAltura = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Piso
        vCampo = "Piso"
        vPosicion = 1
        vPiso = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Depto
        vCampo = "Depto"
        vPosicion = 1
        vDepto = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Torre
        vCampo = "Torre"
        vPosicion = 1
        vTorre = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'CodigoPostal
        vCampo = "CodigoPostal"
        vPosicion = 1
        vgCODIGOPOSTAL = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Barrio
        vCampo = "Barrio"
        vBarrio = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Localidad
        vCampo = "Localidad"
        vPosicion = 1
        vgLOCALIDAD = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'Provincia
        vCampo = "Provincia"
        vPosicion = 1
        vgPROVINCIA = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'IDPRODUCTO
        vCampo = "IDPRODUCTO"
        vPosicion = 2
            v = sLine 'Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
            If Len(v) > 0 Then
            
            'ObtenerCoberturas (vgidCampana), (v), vCantDeErrores, flnErr, Ll, vCampo
            vgCodigoEnCliente = vgIdProducto
            Dim rsprod As New Recordset
            ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, COBERTURAAP, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = '" & vgidCampana & "' and idproductoencliente = '" & v & "'"
            rsprod.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                If Not rsprod.EOF Then
                     vgCOBERTURAVEHICULO = rsprod("COBERTURAVEHICULO")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCOBERTURAHOGAR = rsprod("coberturahogar")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCOBERTURAAP = rsprod("coberturaap")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCodigoEnCliente = v
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                Else
                     vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", Ll, vCampo)

                End If
            rsprod.Close
            End If

'--------------------------------------------------------------------------
    vgAPELLIDOYNOMBRE = Trim(vgApellido) & ", " & Trim(vgNombre)
    vgDOMICILIO = Trim(vCalle) & " " & Trim(vAltura) & " " & Trim(vPiso) & " " & Trim(vDepto) & " " & Trim(vTorre)
    'vgFECHAVIGENCIA = Now
            
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
'       Indicando, en un campo "RegistroRepetido" para no importarlos preo que pueda ser usados
'       para indicar que haria que cambiarle en produccion la corrida y la feca de corrida
'   .
    Dim rscn1 As New Recordset
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where IdCampana = " & vgidCampana & " and Cuit = '" & Trim(vgCuit) & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    
    vgFECHAVENCIMIENTO = DateAdd("m", 12, fechaActual)
    vVigenciaVigente = fechaActual
    'vVigenciaVigente = (rscn1("FECHAVIGENCIA"))
    If Err.Number = 3021 Then 'Limpio error: El valor de BOF o EOF es True, o el actual registro se elimino
        Err.Clear
    End If
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                'If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                'If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
                If Trim(rscn1("Agencia")) <> Trim(vgAgencia) Then vdif = vdif + 1
                If Trim(rscn1("Cargo")) <> Trim(vgCargo) Then vdif = vdif + 1
                If Trim(rscn1("Sexo")) <> Trim(vgSexo) Then vdif = vdif + 1
                'If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                If (Trim(rscn1("Cuit")) <> Trim(vgCuit)) Or IsNull(rscn1("Cuit")) Then vdif = vdif + 1
                'If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAAP")) <> Trim(vgCOBERTURAAP) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If
        
        If vgIDPOLIZA = 0 Then
            vVigenciaVigente = fechaActual
        End If
        
        rscn1.Close
'-=================================================================================================================
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
        ssql = ssql & "IDPOLIZA, "
        ssql = ssql & "IDCIA, "
        'ssql = ssql & "NUMEROCOMPANIA, "
        ssql = ssql & "NROPOLIZA, "
        'ssql = ssql & "NROSECUENCIAL, "
        ssql = ssql & "APELLIDOYNOMBRE, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "CODIGOPOSTAL, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "FECHAVENCIMIENTO, "
        ssql = ssql & "COBERTURAVEHICULO, "
        ssql = ssql & "COBERTURAVIAJERO, "
        ssql = ssql & "COBERTURAHOGAR, "
        ssql = ssql & "COBERTURAAP, "
        'ssql = ssql & "TipodeOperacion, "
        ssql = ssql & "CORRIDA, "
        ssql = ssql & "IdCampana, "
        'ssql = ssql & "TipodeDocumento, "
        ssql = ssql & "NumeroDeDocumento, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "FechaDeNacimiento, "
        ssql = ssql & "InformadoSinCobertura, " 'Agregado
        ssql = ssql & "CodigoEnCliente, "
        ssql = ssql & "Agencia, "
        ssql = ssql & "Cargo, "
        ssql = ssql & "Sexo, "
        ssql = ssql & "Cuit, "
        ssql = ssql & "Modificaciones )"

        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgidCia) & ", '"
        'ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
        ssql = ssql & Trim(vgNROPOLIZA) & "', '"
        'ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgLOCALIDAD) & "', '"
        ssql = ssql & Trim(vgPROVINCIA) & "', '"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
        ssql = ssql & Trim(vVigenciaVigente) & "', '"
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
        ssql = ssql & Trim(vgCOBERTURAAP) & "', "
        'ssql = ssql & Trim(vgTipodeOperacion) & "', "
        ssql = ssql & Trim(vgCORRIDA) & ", "
        ssql = ssql & Trim(vgidCampana) & ", '"
        'ssql = ssql & Trim(vgTipodeDocumento) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        ssql = ssql & Trim(vLote) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vCoberturaFinanciera) & "', '"  'AGREGADO
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
        ssql = ssql & Trim(vgAgencia) & "', '"
        ssql = ssql & Trim(vgCargo) & "', '"
        ssql = ssql & Trim(vgSexo) & "', '"
        ssql = ssql & Trim(vgCuit) & "', '"
        ssql = ssql & Trim(vdif) & "') "
        cn.Execute ssql
        
'========Control de errores=========================================================
        If Err Then
            vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, "")
            Err.Clear
        
        End If
'===========================================================================================

         If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            
        ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & vgidCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
          

            ll100 = 0
        End If
        DoEvents
    Loop


'================Control de Leidos===============================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
    listoParaProcesar
'=================================================================================
        
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos===========================================
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    Dim rsCMP As New Recordset
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " Procesando los datos"
        DoEvents
        rsCMP.Close
    Next lLote

    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lidCampana & ", " & vgCORRIDA
    Procesado
'=====================================================================================
    ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " FinDeProceso"
    ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF


Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub




