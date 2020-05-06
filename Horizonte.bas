Attribute VB_Name = "Horizonte"
Option Explicit
Public Sub ImportarHorizonte()
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
Dim rsCMP As New Recordset
Dim vlineasTotales As Long
Dim vFECHAVIGENCIA As String
Dim vFECHAVENCIMIENTO As String
Dim regMod As Long


'Dim rs As ADODB.Recordset
'    rs.Open Null, Null, adOpenKeyset, adLockBatchOptimistic
'    rs.batchupdate

'On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos523"

On Error Resume Next
vgidCia = lIdCia '10000567
vgidCampana = lIdCampana '523

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)
'======='control de lectura del archivo de datos
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
'=======================================================================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
    

    Ll = 1
    nroLinea = 1
    vLote = 1
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
        
        vgPATENTE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNroMotor = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNROSECUENCIAL = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgTipodeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
       If vgTipodeDocumento = "00" Then
           vgTipodeDocumento = "CI Policia Federal"
       ElseIf vgTipodeDocumento = "80" Then
           vgTipodeDocumento = "CUIT"
       ElseIf vgTipodeDocumento = "89" Then
           vgTipodeDocumento = "LE"
       ElseIf vgTipodeDocumento = "90" Then
           vgTipodeDocumento = "LC"
       ElseIf vgTipodeDocumento = "94" Then
           vgTipodeDocumento = "Pasaporte"
       ElseIf vgTipodeDocumento = "96" Then
           vgTipodeDocumento = "DNI"
       ElseIf vgTipodeDocumento = "99" Then
           vgTipodeDocumento = "CUIL"
       End If
'      -------------------------------------------------------

        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgDOMICILIO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgCODIGOPOSTAL = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgTelefono = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgMARCADEVEHICULO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgAno = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vFECHAVIGENCIA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        vgFECHAVIGENCIA = Mid(vFECHAVIGENCIA, 7, 2) & "/" & Mid(vFECHAVIGENCIA, 5, 2) & "/" & Mid(vFECHAVIGENCIA, 1, 4)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vFECHAVENCIMIENTO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        vgFECHAVENCIMIENTO = Mid(vFECHAVENCIMIENTO, 7, 2) & "/" & Mid(vFECHAVENCIMIENTO, 5, 2) & "/" & Mid(vFECHAVENCIMIENTO, 1, 4)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'        vgtipoRC = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
'        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgCOBERTURAVEHICULO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgOperacion = sLine

        Select Case vgPROVINCIA
            Case "01"
            vgPROVINCIA = "CAPITAL"
            Case "02"
            vgPROVINCIA = "BUENOS AIRES"
            Case "03"
            vgPROVINCIA = "CATAMARCA"
            Case "04"
            vgPROVINCIA = "CORDOBA"
            Case "05"
            vgPROVINCIA = "CORRIENTES"
            Case "06"
            vgPROVINCIA = "ENTRE RIOS"
            Case "07"
            vgPROVINCIA = "JUJUY"
            Case "08"
            vgPROVINCIA = "LA RIOJA"
            Case "09"
            vgPROVINCIA = "MENDOZA"
            Case "10"
            vgPROVINCIA = "SALTA"
            Case "11"
            vgPROVINCIA = "SAN JUAN"
            Case "12"
            vgPROVINCIA = "SAN LUIS"
            Case "13"
            vgPROVINCIA = "SANTA FE"
            Case "14"
            vgPROVINCIA = "SANTIAGO DEL ESTERO"
            Case "15"
            vgPROVINCIA = "TUCUMAN"
            Case "16"
            vgPROVINCIA = "CHACO"
            Case "17"
            vgPROVINCIA = "CHUBUT"
            Case "18"
            vgPROVINCIA = "FORMOSA"
            Case "19"
            vgPROVINCIA = "LA PAMPA"
            Case "21"
            vgPROVINCIA = "MISIONES"
            Case "22"
            vgPROVINCIA = "Neuquen"
            Case "23"
            vgPROVINCIA = "Rio Negro"
            Case "24"
            vgPROVINCIA = "Santa Cruz"
            Case "25"
            vgPROVINCIA = "TierraDelFuego"
            Case "26"
            vgPROVINCIA = "Exterior"
        End Select
 
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
                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
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
                If Trim(rscn1("Gama")) <> Trim(vgGama) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If

        rscn1.Close
'-=================================================================================================================
 
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos523 ("
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
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones )"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', "
            ssql = ssql & Trim(vgIDAUTO) & ", '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgCOLOR) & "', '"
            ssql = ssql & Trim(vgAno) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', "
            ssql = ssql & Trim(vgTIPODEVEHICULO) & ", '"
            ssql = ssql & Trim(vgTipodeServicio) & "', '"
            ssql = ssql & Trim(vgIDTIPODECOBERTURA) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgOperacion) & "', '"
            ssql = ssql & Trim(vgCATEGORIA) & "', '"
            ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", "
            ssql = ssql & Trim(vgidCampana) & ", '"
            ssql = ssql & Trim(vgConductor) & "', '"
            ssql = ssql & Trim(vgCodigoDeProductor) & "', '"
            ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeHogar) & "', '"
            ssql = ssql & Trim(vgIniciodeAnualidad) & "', '"
            ssql = ssql & Trim(vgPolizaIniciaAnualidad) & "', '"
            ssql = ssql & Trim(vgTelefono) & "', '"
            ssql = ssql & Trim(vgNroMotor) & "', '"
            ssql = ssql & Trim(vgGama) & "', '"
            ssql = ssql & Trim(vLote) & "', '"
            ssql = ssql & Trim(vdif) & "') "
            cn.Execute ssql
'        End If
        
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
    vlineasTotales = Ll
'    Ll = 0
'    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
'    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
'    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
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
    
'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesadosSoloNovedades " & lIdCampana & ", " & vgCORRIDA
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




