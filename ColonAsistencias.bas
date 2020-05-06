Attribute VB_Name = "ColonAsistencias"
Public Sub ImportarColonAsistencias()

Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim ssqlProv As String
Dim rsCMPProv As New Recordset
Dim rsp As New Recordset
Dim rsprod As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim vlineasTotales As Long
Dim sArchivo As String
Dim vHoja As String
Dim lRowH As Long
Dim vtotalDeLineas As Long
Dim vdir As Integer
Dim regMod As Long
Dim vVigenciaVigente As String
Dim fechaActual As Date

'Variables para Colon Seguros
Dim vDepto As String
Dim vCalle As String
Dim vAltura As String
Dim vApellido As String
Dim vNombre As String
Dim vNro As String
Dim vPiso As String
Dim Ll As Long
Dim vFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim vLinea As Long
Dim vPosicion As Long
Dim vCampo As String

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

fechaActual = Now

TablaTemporal

On Error Resume Next

'================
Dim vCS As String
vCS = ";"
'================

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

Ll = 2
ll100 = 0
'vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & "Chubb.txt"
vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
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
'cn1.Execute "TM_CargaPolizasLogDeSetCorridasxcia " & lIdCia & ", " & vgCORRIDA
cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lidCampana & ", " & vgCORRIDA
'ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcia = " & lIdCia & " and Registrosleidos is null"
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
Dim vControlDeModificados As Long
LongDeLote = 1000
nroLinea = 1
vLote = 1
'=============================================================================================

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If
tf.SkipLine 'Saltea encabezado
Do Until tf.AtEndOfStream
    vLinea = Ll
    sLine = tf.ReadLine
    If Len(Trim(sLine)) < 5 Then Exit Do
    sLine = Replace(sLine, "'", "")
        
    '====maneja los lotes para corte de importacion========
    nroLinea = nroLinea + 1
    If nroLinea = LongDeLote + 1 Then
        vLote = vLote + 1
        vControlDeModificados = 0
        nroLinea = 1
    End If
    '======================================================
    Blanquear
    vPosicion = 0
'==================================================================================================
    vCampo = "Producto"
    vPosicion = vPosicion + 1
        'vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Operacion"
    vPosicion = vPosicion + 1
        vgOperacion = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "ID"
    vPosicion = vPosicion + 1
        'vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Id.Cliente"
    vPosicion = vPosicion + 1
        'vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Cód.Asistencia"
    vPosicion = vPosicion + 1
        v = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        If Len(v) > 0 Then
            ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente = '" & v & "'"
            rsprod.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            If Not rsprod.EOF Then
                vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                vgCOBERTURAHOGAR = rsprod("coberturahogar")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                vgCodigoEnCliente = v
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
            Else
                vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", Ll, vCampo)
            End If
            rsprod.Close
        End If
    'vgCodigoEnCliente = sLine 'Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Fec.Alta.Asist"
    vPosicion = vPosicion + 1
        vgFECHAVIGENCIA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Apellido y Nombre/Razón Social"
    vPosicion = vPosicion + 1
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Id.Tipo Doc"
    vPosicion = vPosicion + 1
        vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Nro.Doc"
    vPosicion = vPosicion + 1
        vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Fecha nacimiento"
    vPosicion = vPosicion + 1
        'vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Sexo"
    vPosicion = vPosicion + 1
        vgSexo = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        If vgSexo = "FEMENINO" Then
            vgSexo = "F"
        ElseIf vgSexo = "MASCULINO" Then
            vgSexo = "M"
        End If
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Email"
    vPosicion = vPosicion + 1
        vgEmail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Calle"
    vPosicion = vPosicion + 1
        vCalle = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Altura"
    vPosicion = vPosicion + 1
        vNro = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Piso"
    vPosicion = vPosicion + 1
        vPiso = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Dpto"
    vPosicion = vPosicion + 1
        vDepto = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Localidad"
    vPosicion = vPosicion + 1
        vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "CP"
    vPosicion = vPosicion + 1
        vgCODIGOPOSTAL = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Provincia"
    vPosicion = vPosicion + 1
        vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Pais"
    vPosicion = vPosicion + 1
        vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Telefono1"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Telefono2"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Telefono3"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Sponsor" ' empresa o agencia, elegi Conductor pq ACE lo usa
    vPosicion = vPosicion + 1
        vgConductor = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Id.Tip.Benef"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Id.Tip.Asist"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Id.Compañia"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Adicional1"
    vPosicion = vPosicion + 1
        vgPATENTE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Adicional2"
    vPosicion = vPosicion + 1
        vgMODELO = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Adicional3"
    vPosicion = vPosicion + 1
        vgAno = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Adicional4"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
    vCampo = "Adicional5"
    vPosicion = vPosicion + 1
        'vgPais = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
    sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
   
    vgDOMICILIO = vCalle & " " & vNro & " " & vPiso & " " & vDepto
    'vgAPELLIDOYNOMBRE = vApellido & " " & vNombre
    
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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and NumeroDeDocumento = '" & Trim(vgNumeroDeDocumento) & "' and TipodeDocumento = '" & Trim(vgTipodeDocumento) & "'" 'and TipodeServicio = '" & vgTipodeServicio & "' and CodigoDeProductor = '" & vgCodigoDeProductor & "'"
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    
    vgFECHAVENCIMIENTO = DateAdd("m", 2, fechaActual)
    'vVigenciaVigente = (rscn1("FECHAVIGENCIA"))

    vgIDPOLIZA = 0
    If Not rscn1.EOF Then
        vdif = 0  'setea la variale de control de repetido con modificacion en cero
        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
        If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
        If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
        If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
        If Trim(rscn1("CodigoEnCliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
        If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
        If vgCOBERTURAHOGAR <> "" Then
            If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
        End If
        If vgCOBERTURAVEHICULO <> "" Then
            If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
        End If
        If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
        If Trim(rscn1("OPERACION")) <> Trim(vgOperacion) Then vdif = vdif + 1
        If Trim(rscn1("EMAIL")) <> Trim(vgEmail) Then vdif = vdif + 1
        If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
        If Trim(rscn1("PAIS")) <> Trim(vgPais) Then vdif = vdif + 1
        If Trim(rscn1("CONDUCTOR")) <> Trim(vgConductor) Then vdif = vdif + 1
        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
        If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
        vgIDPOLIZA = rscn1("idpoliza")
    End If

    rscn1.Close
'-=================================================================================================================
    ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & lidCampana & "("
    ssql = ssql & "IDPOLIZA, "
    ssql = ssql & "IDCIA, "
    ssql = ssql & "IdCampana, "
    ssql = ssql & "NROPOLIZA, "
    ssql = ssql & "APELLIDOYNOMBRE, "
    ssql = ssql & "DOMICILIO, "
    ssql = ssql & "LOCALIDAD, "
    ssql = ssql & "CODIGOPOSTAL, "
    ssql = ssql & "Sexo, "
    ssql = ssql & "FECHAVIGENCIA, "
    ssql = ssql & "FECHAVENCIMIENTO, "
    ssql = ssql & "COBERTURAHOGAR, "
    ssql = ssql & "COBERTURAVEHICULO, "
    ssql = ssql & "COBERTURAVIAJERO, "
    ssql = ssql & "CORRIDA, "
    ssql = ssql & "TipodeDocumento, "
    ssql = ssql & "NumeroDeDocumento, "
    ssql = ssql & "Documento, "
    ssql = ssql & "Operacion, "
    ssql = ssql & "Email, "
    ssql = ssql & "Provincia, "
    ssql = ssql & "Pais, "
    ssql = ssql & "Conductor, "
    ssql = ssql & "Patente, "
    ssql = ssql & "Modelo, "
    ssql = ssql & "Ano, "
    ssql = ssql & "CodigoEnCliente, "
    ssql = ssql & "IdLote, "
    ssql = ssql & "Modificaciones)"

    ssql = ssql & " values( "
    ssql = ssql & Trim(vgIDPOLIZA) & ", "
    ssql = ssql & Trim(vgidCia) & ", "
    ssql = ssql & Trim(vgidCampana) & ", '"
    ssql = ssql & Trim(vgNROPOLIZA) & "', '"
    ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
    ssql = ssql & Trim(vgDOMICILIO) & "', '"
    ssql = ssql & Trim(vgLOCALIDAD) & "', '"
    ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
    ssql = ssql & Trim(vgSexo) & "', '"
    ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
    ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
    ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
    ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
    ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
    ssql = ssql & Trim(vgCORRIDA) & "', '"
    ssql = ssql & Trim(vgTipodeDocumento) & "', '"
    ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
    ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
    ssql = ssql & Trim(vgOperacion) & "', '"
    ssql = ssql & Trim(vgEmail) & "', '"
    ssql = ssql & Trim(vgPROVINCIA) & "', '"
    ssql = ssql & Trim(vgPais) & "', '"
    ssql = ssql & Trim(vgConductor) & "', '"
    ssql = ssql & Trim(vgPATENTE) & "', '"
    ssql = ssql & Trim(vgMODELO) & "', '"
    ssql = ssql & Trim(vgAno) & "', '"
    ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
    ssql = ssql & Trim(vLote) & "', '"
    ssql = ssql & Trim(vdif) & "') "
    cn.Execute ssql
    
'========Control de errores=========================================================
    If Err Then
        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, vCampo)
        'MsgBox "PARA"
        Err.Clear
    
    End If
'====================================================================================

        If vdif > 0 Then
        regMod = regMod + 1
    End If
    
    Ll = Ll + 1
    ll100 = ll100 + 1
    If ll100 = 100 Then
        ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
    ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

            ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lidCampana & " and corrida =" & vgCORRIDA
            cn1.Execute ssql
        ll100 = 0
    End If
    DoEvents
Loop

'================Control de Leidos================================================
cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
listoParaProcesar
'=================================================================================
ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos====================================
cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'=================================================================================
ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
Dim rsCMP As New Recordset
DoEvents
For lLote = 1 To vLote
    cn1.CommandTimeout = 300
    cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & lidCampana
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
