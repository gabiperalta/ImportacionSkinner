Attribute VB_Name = "BeneficioSeguros"
Public Sub ImportarBeneficioSeguros()

Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim rsprod As New Recordset
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim v
Dim vCampo As String
Dim vPosicion As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
Dim vlineasTotales As Long
Dim vSinImportar As String
Dim vCoberturaFinanciera As String
Dim vFechaDeNacimiento As String
Dim vgFechaNacimiento As String
Dim vgVNumeroPoliza As String
Dim vFechaDeVigencia As String
Dim vFechaDeVencimiento As String
Dim vVigenciaVigente As String
Dim fechaActual As Date

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lIdCampana
fechaActual = Now
'======================
TablaTemporal
'======================
Dim vCS As String
vCS = ";"
'======================
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

    'Ll = 0
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
    cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lIdCampana & ", " & vgCORRIDA
    ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcampana = " & lIdCampana & " and Registrosleidos is null"
    rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    If rsCorr.EOF Then
        MsgBox "no se determino la corrida, se detendra el proceso"
        Exit Sub
    Else
        vgCORRIDA = rsCorr("corrida")
    End If
'=======seteo control de lote==========================================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
'======================================================================
    Ll = 2
    tf.SkipLine 'Saltea la linea de encabezados
    Do Until tf.AtEndOfStream
        Blanquear
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
        vPosicion = 0
'NOMBRE
        vCampo = "APELLIDOYNOMBRE"
        vPosicion = vPosicion + 1
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'RAMA
        vCampo = "RAMA"
        vPosicion = vPosicion + 1
        vgOBSERVACIONES = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'NROPOLIZA
        vCampo = "NROPOLIZA"
        vPosicion = vPosicion + 1
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'DOMICILIO
        vCampo = "DOMICILIO"
        vPosicion = vPosicion + 1
        vgDOMICILIO = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'TELEFONO
        vCampo = "TELEFONO"
        vPosicion = vPosicion + 1
        vgTelefono = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'FECHANACIMIENTO
        vCampo = "FECHANACIMIENTO"
        vPosicion = vPosicion + 1
        vgFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        If IsEmpty(vgFechaDeNacimiento) Or Not IsDate(vgFechaDeNacimiento) Then
            vgFechaDeNacimiento = Now
        End If
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'PRODUCTO
        vCampo = "PRODUCTO"
        vPosicion = vPosicion + 1
        v = sLine
        If Len(v) > 0 Then
            ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lIdCampana & "  and idproductoencliente = " & v
            rsprod.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                If Not rsprod.EOF Then
                     vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, sName)
                     vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, sName)
                     vgCOBERTURAHOGAR = rsprod("coberturahogar")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, sName)
                     vgCodigoEnCliente = v
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, sName)
                Else
                     vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", lRow, sName)
                
                End If
            rsprod.Close
        End If
        'Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
            
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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where IdCampana = " & vgidCampana & " and NROPOLIZA = '" & vgNROPOLIZA & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    
    vgFECHAVENCIMIENTO = DateAdd("m", 3, fechaActual)
    vVigenciaVigente = (rscn1("FECHAVIGENCIA"))
    
    If Err.Number = 3021 Then 'Limpio error: El valor de BOF o EOF es True, o el actual registro se elimino
        Err.Clear
    End If
    
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                If Trim(rscn1("OBSERVACIONES")) <> Trim(vgOBSERVACIONES) Then vdif = vdif + 1
                If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                If Trim(rscn1("CodigoEnCliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
                If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
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
        ssql = ssql & "CodigoEnCliente, "
        ssql = ssql & "NROPOLIZA, "
        ssql = ssql & "APELLIDOYNOMBRE, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "Telefono, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "FECHAVENCIMIENTO, "
        ssql = ssql & "OBSERVACIONES, "
        ssql = ssql & "NumeroDeDocumento, "
        ssql = ssql & "FechaDeNacimiento, "
        ssql = ssql & "COBERTURAVEHICULO, "
        ssql = ssql & "COBERTURAVIAJERO, "
        ssql = ssql & "COBERTURAHOGAR, "
        ssql = ssql & "CORRIDA, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "Modificaciones)"
        
        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgidCia) & ", '"
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
        ssql = ssql & Trim(vgNROPOLIZA) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgTelefono) & "', '"
        ssql = ssql & Trim(vVigenciaVigente) & "', '"
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgOBSERVACIONES) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', "
        ssql = ssql & Trim(vgCORRIDA) & ", '"
        ssql = ssql & Trim(vLote) & "', '"
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
    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
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

