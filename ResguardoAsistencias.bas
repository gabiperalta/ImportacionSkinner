Attribute VB_Name = "ResguardoAsistencias"
Option Explicit
Public Sub ImportarResguardoAsistencias()
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
'Dim nroLinea As Long
Dim vCampo As String
Dim vPosicion As Long
Dim regMod As Long
'Dim lLote As Long
'Dim vLote As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
'Dim vIDCIA As Long
'Dim vIDCampana As Long
'Dim rsCMP As New Recordset
'Dim LongDeLote As Integer
'Dim vlineasTotales As Long
'Dim rs As ADODB.Recordset
'    rs.Open Null, Null, adOpenKeyset, adLockBatchOptimistic
'    rs.batchupdate

cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos524"

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

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
        rsCorr.Close
        
        ' se obtiene el idHistorialDeImportacion
        ssql = "SELECT IdHistorialDeImportacion "
        ssql = ssql & " FROM tm_ImportacionHistorial "
        ssql = ssql & " WHERE corrida = " & vgCORRIDA & " and idCampana = " & vgidCampana
        
        rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        vgIdHistorialImportacion = rsCorr("IdHistorialDeImportacion")
        rsCorr.Close
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
    nroLinea = 1
    vLote = 1
    
    InicializarCoberturaLista
    
    Do Until tf.AtEndOfStream
        vCoberturaActual(0) = ""
        vCoberturaActual(1) = ""
        vCoberturaActual(2) = ""
    
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


 '      -------------------------------------------------------
        vCampo = "NroPoliza"
        vPosicion = 30
        vgNROPOLIZA = Mid(Trim(Mid(sLine, 30, 80)), 1, 20)
 '      -------------------------------------------------------
        vCampo = "APELLIDO"
        vPosicion = 110
        vgAPELLIDOYNOMBRE = Trim(Mid(sLine, 110, 40))
 '      -------------------------------------------------------
        vCampo = "Nombre"
        vPosicion = 150
        vgAPELLIDOYNOMBRE = Trim(vgAPELLIDOYNOMBRE) & "," & Trim(Mid(sLine, 150, 40))
 '      -------------------------------------------------------
        vCampo = "DOMICILIO"
        vPosicion = 211
        vgDOMICILIO = Mid(sLine, 211, 80)
 '      -------------------------------------------------------
        vCampo = "LOCALIDAD"
        vPosicion = 341
        vgLOCALIDAD = Mid(sLine, 341, 70)
 '      -------------------------------------------------------
        vCampo = "PROVINCIA"
        vPosicion = 431
        vgPROVINCIA = Mid(sLine, 431, 2)
        Select Case vgPROVINCIA
                Case "B"
                    vgPROVINCIA = "Buenos Aires"
                Case "C"
                    vgPROVINCIA = "Capital Federal"
                Case "x"
                    vgPROVINCIA = "Cordoba"
        End Select
 '      -------------------------------------------------------
        vCampo = "CODIGOPOSTAL"
        vPosicion = 301
        vgCODIGOPOSTAL = Mid(sLine, 301, 8)
 '      -------------------------------------------------------
        vCampo = "FECHAVIGENCIA"
        vPosicion = 542
        vgFECHAVIGENCIA = Mid(sLine, 526, 2) & "/" & Mid(sLine, 529, 2) & "/" & Mid(sLine, 532, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 549
        vgFECHAVENCIMIENTO = Mid(sLine, 536, 2) & "/" & Mid(sLine, 539, 2) & "/" & Mid(sLine, 542, 4)
 '      -------------------------------------------------------
        vCampo = "MARCADEVEHICULO"
        vPosicion = 473
        vgMARCADEVEHICULO = Mid(sLine, 473, 35)
 '      -------------------------------------------------------
        vCampo = "COLOR"
        vPosicion = 512
        vgCOLOR = Mid(sLine, 512, 10)
 '      -------------------------------------------------------
        vCampo = "AÑO"
        vPosicion = 508
        vgAno = Mid(sLine, 508, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 1
        vgPATENTE = Mid(sLine, 1, 9)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVEHICULO"
        vPosicion = 524
        vgCOBERTURAVEHICULO = Mid(sLine, 524, 2)
 '      -------------------------------------------------------
        vCampo = "Telefono"
        vPosicion = 433
        vgTelefono = Mid(Trim(Mid(sLine, 433, 40)), 1, 10)
 '      -------------------------------------------------------
 
        '=============== Lectura de coberturas ===============
        LeerCoberturas
        
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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' and Nrosecuencial = '" & vgNROSECUENCIAL & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
                If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
'                If Trim(rscn1("IDTIPODECOBERTURA")) <> Trim(vgIDTIPODECOBERTURA) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
'                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then vdif = vdif + 1
                If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then vdif = vdif + 1
                If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then vdif = vdif + 1
'                If Trim(rscn1("IdCampana")) <> Trim(vgIdCampana) Then vdif = vdif + 1
                If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then vdif = vdif + 1
                If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then vdif = vdif + 1
                If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then vdif = vdif + 1
                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
'                If Trim(rscn1("Referido")) <> Trim(vgReferido) Then vdif = vdif + 1
                If Trim(rscn1("Gama")) <> Trim(vgGama) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If

        rscn1.Close
'-=================================================================================================================
 

            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos524 ("
            ssql = ssql & "IDPOLIZA, "
            ssql = ssql & "IDCIA, "
            ssql = ssql & "NUMEROCOMPANIA, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "NROSECUENCIAL, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
            ssql = ssql & "CODIGOPOSTAL, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "IDAUTO, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "COLOR, "
            ssql = ssql & "ANO, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "TIPODEVEHICULO, "
            ssql = ssql & "TipodeServicio, "
            ssql = ssql & "IDTIPODECOBERTURA, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "TipodeOperacion, "
            ssql = ssql & "Operacion, "
            ssql = ssql & "CATEGORIA, "
            ssql = ssql & "ASISTENCIAXENFERMEDAD, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "Conductor, "
            ssql = ssql & "CodigoDeProductor, "
            ssql = ssql & "CodigoDeServicioVip, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "TipodeHogar, "
            ssql = ssql & "IniciodeAnualidad, "
            ssql = ssql & "PolizaIniciaAnualidad, "
            ssql = ssql & "Telefono, "
            ssql = ssql & "NroMotor, "
            ssql = ssql & "Gama, "
            ssql = ssql & "IdProducto, "
            ssql = ssql & "coberturahogar, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones )"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '" 'AAA
            ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '" 'AA
            ssql = ssql & Trim(vgDOMICILIO) & "', '" 'AA
            ssql = ssql & Trim(vgLOCALIDAD) & "', '" 'AA
            ssql = ssql & Trim(vgPROVINCIA) & "', '" ' AA
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '" 'AA
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '" 'AA
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', " 'AA
            ssql = ssql & Trim(vgIDAUTO) & ", '" 'ESTE
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '" 'AA
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgCOLOR) & "', '" 'AA
            ssql = ssql & Trim(vgAno) & "', '" 'AA
            ssql = ssql & Trim(vgPATENTE) & "', " ' AA
            ssql = ssql & Trim(vgTIPODEVEHICULO) & ", '" 'ESTE
            ssql = ssql & Trim(vgTipodeServicio) & "', '"
            ssql = ssql & Trim(vgIDTIPODECOBERTURA) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '" ' AA
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgOperacion) & "', '"
            ssql = ssql & Trim(vgCATEGORIA) & "', '"
            ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", " 'ESET
            ssql = ssql & Trim(vgidCampana) & ", '" 'ESTE
            ssql = ssql & Trim(vgConductor) & "', '"
            ssql = ssql & Trim(vgCodigoDeProductor) & "', '"
            ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeHogar) & "', '"
            ssql = ssql & Trim(vgIniciodeAnualidad) & "', '"
            ssql = ssql & Trim(vgPolizaIniciaAnualidad) & "', '"
            ssql = ssql & Trim(vgTelefono) & "', '" 'AA
            ssql = ssql & Trim(vgNroMotor) & "', '"
            ssql = ssql & Trim(vgGama) & "', '"
            ssql = ssql & Trim(vgIdProducto) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
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

'========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lidCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql

            ll100 = 0
        End If
        DoEvents
    Loop
    
    ProcesarCoberturasLeidas
    
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
    CoberturasProcesadas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
    CoberturasBajas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
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




