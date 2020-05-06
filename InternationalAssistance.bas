Attribute VB_Name = "InternationalAssistance"
Option Explicit
Public Sub ImportarInternationalAssistance()
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
Dim vlineasTotales As Long
Dim vgDia As Integer
Dim vgMes As Integer
Dim vgAno As Integer
Dim regMod As Long
Dim vdif As Integer

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lIdCampana

TablaTemporal

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
    vdif = 0
    Ll = 1
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

        vCampo = "Compania"
        vPosicion = 1
        vgNUMEROCOMPANIA = Mid(sLine, 1, 3)
 '      -------------------------------------------------------
        vCampo = "NroPoliza"
        vPosicion = 4
        vgNROPOLIZA = Mid(sLine, 4, 20)
 '      -------------------------------------------------------
        vCampo = "NroSecuencial"
        vPosicion = 17
        vgNROSECUENCIAL = Mid(sLine, 24, 3)
 '      -------------------------------------------------------
        vCampo = "APELLIDOYNOMBRE"
        vPosicion = 20
        vgAPELLIDOYNOMBRE = Mid(sLine, 27, 50)
 '      -------------------------------------------------------
        vCampo = "DOMICILIO"
        vPosicion = 70
        vgDOMICILIO = Mid(sLine, 77, 55)
 '      -------------------------------------------------------
'        vCampo = "LOCALIDAD"
'        vPosicion = 125
'        vgLOCALIDAD = Mid(sLine, 132, 30)
' '      -------------------------------------------------------
'        vCampo = "PROVINCIA"
'        vPosicion = 155
'        vgPROVINCIA = Mid(sLine, 155, 15)
 '      -------------------------------------------------------
        vCampo = "CODIGOPOSTAL"
        vPosicion = 170
        vgCODIGOPOSTAL = Mid(sLine, 132, 8)
 '      -------------------------------------------------------
        vCampo = "FECHAVIGENCIA"
        vPosicion = 184
        vgFECHAVIGENCIA = Mid(sLine, 146, 2) & "/" & Mid(sLine, 144, 2) & "/" & Mid(sLine, 140, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 192
        vgFECHAVENCIMIENTO = Mid(sLine, 154, 2) & "/" & Mid(sLine, 152, 2) & "/" & Mid(sLine, 148, 4)
 '      -------------------------------------------------------
'        vgFECHAALTAOMNIA = Mid(sLine, 1, 10)
'        vgFECHABAJAOMNIA = Mid(sLine, 1, 10)
'        vgIDAUTO = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "MARCADEVEHICULO"
        vPosicion = 194
        vgMARCADEVEHICULO = Mid(sLine, 156, 30)
 '      -------------------------------------------------------
        vCampo = "MODELO"
        vPosicion = 224
        vgMODELO = Mid(sLine, 186, 20)
 '      -------------------------------------------------------
        vCampo = "COLOR"
        vPosicion = 244
        vgCOLOR = Mid(sLine, 206, 30)
 '      -------------------------------------------------------
        vCampo = "AÑO"
        vPosicion = 259
        vgAno = Mid(sLine, 236, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 263
        vgPATENTE = Mid(sLine, 240, 15)
        vgNROPOLIZA = vgPATENTE
 '      -------------------------------------------------------
'        vCampo = "TipodeServicio"
'        vPosicion = 271
'        vgTipodeServicio = Mid(sLine, 271, 4)
 '      -------------------------------------------------------
 '       vCampo = "TIPODEVEHICULO"
 '       vPosicion = 275
 '       If IsNumeric(Mid(sLine, 275, 2)) Then
 '           vgTIPODEVEHICULO = Mid(sLine, 275, 2)
 '       Else
 '           vgTIPODEVEHICULO = 0
 '       End If
 '      -------------------------------------------------------
'        vgIDTIPODECOBERTURA = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVEHICULO"
        vPosicion = 277
        vgCOBERTURAVEHICULO = Mid(sLine, 257, 1)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVIAJERO"
        vPosicion = 279
        vgCOBERTURAVIAJERO = Mid(sLine, 258, 1)
        
 '      -------------------------------------------------------
        vCampo = "COBERTURAHOGAR"
        vPosicion = 279
        vgCOBERTURAHOGAR = Mid(sLine, 259, 1)
 '      -------------------------------------------------------

        vgCodigoDeProceso = 0
        vgCodigoDeProceso = Mid(sLine, 271, 1)
        
'=========Correccion iveco==============================
        If UCase(Trim(vgMARCADEVEHICULO)) = "IVECO" Then
            vgCOBERTURAVEHICULO = "01"
            vgCOBERTURAVIAJERO = "04"
            vgTIPODEVEHICULO = 4
        End If
'=======================================================
        
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
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
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "TipodeOperacion, "
            ssql = ssql & "Operacion, "
            ssql = ssql & "CATEGORIA, "
            ssql = ssql & "ASISTENCIAXENFERMEDAD, "
            ssql = ssql & "CORRIDA, "
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
            ssql = ssql & "CodigoDeProceso, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vgidCia) & ", "
            ssql = ssql & Trim(vgidCampana) & ", '"
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
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgOperacion) & "', '"
            ssql = ssql & Trim(vgCATEGORIA) & "', '"
            ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", '"
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
            ssql = ssql & Trim(vgCodigoDeProceso) & "', '"
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
        cn1.Execute sSPImportacion & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " Procesando los datos"
        DoEvents
        rsCMP.Close
    Next lLote
    
'    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
'!!!!Esta compania informa altas y modificaciones, no hay que dar de bajas a las que no informa!!!!
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


Public Sub ImportarInternationalAssistanceOld()
'falta ingresar la campaña en el reporte de importacion, hay error al poner a procesar el archivo ( el archivo no se modifica).
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim nroLinea As Long
Dim vCampo As String
Dim vPosicion As Long
Dim lLote As Long
Dim vLote As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
Dim vIDCIA As Long
Dim vIDCampana As Long
Dim rsCMP As New Recordset
Dim LongDeLote As Integer
Dim vlineasTotales As Long
On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosv2 where idcia=" & 9999652 & "and idcampana= " & 330
    vIDCIA = 9999652
    vIDCampana = 330
    LongDeLote = 1000

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)
    Ll = 1
    nroLinea = 1
    vLote = 1
    Do Until tf.AtEndOfStream
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
        
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
        
'("NUMEROCOMPANIA") = DTSSource("Col001")
'("NROPOLIZA") = DTSSource("Col002")
'("NROSECUENCIAL") = DTSSource("Col003")
'("APELLIDOYNOMBRE") = DTSSource("Col004")
'("DOMICILIO") = DTSSource("Col005")
'("LOCALIDAD") = DTSSource("Col006")
'("PROVINCIA") = DTSSource("Col007")
'("CODIGOPOSTAL") = DTSSource("Col008")
'("FECHAVIGENCIA") = DateSerial(  DTSSource("Col009") , DTSSource("Col010") , DTSSource("Col011")  )
'("FECHAVENCIMIENTO") = DateSerial( DTSSource("Col012") , DTSSource("Col013")  , DTSSource("Col014")
'("MARCADEVEHICULO") = DTSSource("Col015")
'("MODELO") = DTSSource("Col016")
'("COLOR") = DTSSource("Col017")
'("ANO") = DTSSource("Col018")
'("PATENTE") = DTSSource("Col019")
'("TipodeServicio") = DTSSource("Col021")
'("COBERTURAVEHICULO") = DTSSource("Col022")
'("COBERTURAVIAJERO") = DTSSource("Col023")
'("TipodeOperacion") = DTSSource("Col024")
'("Conductor") = DTSSource("Col025")
'("CodigoDeProductor") = DTSSource("Col026")
'("CodigoDeServicioVip") = DTSSource("Col027")

'        vgIDPOLIZA = Mid(sLine, 1, 10)
'        vgIDCIA = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "Compania"
        vPosicion = 1
        vgNUMEROCOMPANIA = Mid(sLine, 1, 3)
 '      -------------------------------------------------------
        vCampo = "NroPoliza"
        vPosicion = 4
        vgNROPOLIZA = Mid(sLine, 4, 20)
 '      -------------------------------------------------------
        vCampo = "NroSecuencial"
        vPosicion = 17
        vgNROSECUENCIAL = Mid(sLine, 24, 3)
 '      -------------------------------------------------------
        vCampo = "APELLIDOYNOMBRE"
        vPosicion = 20
        vgAPELLIDOYNOMBRE = Mid(sLine, 27, 50)
 '      -------------------------------------------------------
        vCampo = "DOMICILIO"
        vPosicion = 70
        vgDOMICILIO = Mid(sLine, 77, 55)
 '      -------------------------------------------------------
'        vCampo = "LOCALIDAD"
'        vPosicion = 125
'        vgLOCALIDAD = Mid(sLine, 132, 30)
' '      -------------------------------------------------------
'        vCampo = "PROVINCIA"
'        vPosicion = 155
'        vgPROVINCIA = Mid(sLine, 155, 15)
 '      -------------------------------------------------------
        vCampo = "CODIGOPOSTAL"
        vPosicion = 170
        vgCODIGOPOSTAL = Mid(sLine, 132, 8)
 '      -------------------------------------------------------
        vCampo = "FECHAVIGENCIA"
        vPosicion = 184
        vgFECHAVIGENCIA = Mid(sLine, 146, 2) & "/" & Mid(sLine, 144, 2) & "/" & Mid(sLine, 140, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 192
        vgFECHAVENCIMIENTO = Mid(sLine, 154, 2) & "/" & Mid(sLine, 152, 2) & "/" & Mid(sLine, 148, 4)
 '      -------------------------------------------------------
'        vgFECHAALTAOMNIA = Mid(sLine, 1, 10)
'        vgFECHABAJAOMNIA = Mid(sLine, 1, 10)
'        vgIDAUTO = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "MARCADEVEHICULO"
        vPosicion = 194
        vgMARCADEVEHICULO = Mid(sLine, 156, 30)
 '      -------------------------------------------------------
        vCampo = "MODELO"
        vPosicion = 224
        vgMODELO = Mid(sLine, 186, 20)
 '      -------------------------------------------------------
        vCampo = "COLOR"
        vPosicion = 244
        vgCOLOR = Mid(sLine, 206, 30)
 '      -------------------------------------------------------
        vCampo = "AÑO"
        vPosicion = 259
        vgAno = Mid(sLine, 236, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 263
        vgPATENTE = Mid(sLine, 240, 15)
        vgNROPOLIZA = vgPATENTE
 '      -------------------------------------------------------
'        vCampo = "TipodeServicio"
'        vPosicion = 271
'        vgTipodeServicio = Mid(sLine, 271, 4)
 '      -------------------------------------------------------
 '       vCampo = "TIPODEVEHICULO"
 '       vPosicion = 275
 '       If IsNumeric(Mid(sLine, 275, 2)) Then
 '           vgTIPODEVEHICULO = Mid(sLine, 275, 2)
 '       Else
 '           vgTIPODEVEHICULO = 0
 '       End If
 '      -------------------------------------------------------
'        vgIDTIPODECOBERTURA = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVEHICULO"
        vPosicion = 277
        vgCOBERTURAVEHICULO = Mid(sLine, 257, 1)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVIAJERO"
        vPosicion = 279
        vgCOBERTURAVIAJERO = Mid(sLine, 258, 1)
        
 '      -------------------------------------------------------
        vCampo = "COBERTURAHOGAR"
        vPosicion = 279
        vgCOBERTURAHOGAR = Mid(sLine, 259, 1)
 '      -------------------------------------------------------

        vgCodigoDeProceso = 0
        vgCodigoDeProceso = Mid(sLine, 271, 1)
        
        
'=========Correccion iveco==============================
        If UCase(Trim(vgMARCADEVEHICULO)) = "IVECO" Then
            vgCOBERTURAVEHICULO = "01"
            vgCOBERTURAVIAJERO = "04"
            vgTIPODEVEHICULO = 4
        End If

'=======================================================
        
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatosv2 ("
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
            ssql = ssql & "COBERTURAHOGAR, "
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
            ssql = ssql & "CodigoDeProceso, "
            ssql = ssql & "IdLote )"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vIDCIA) & ", '"
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
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgOperacion) & "', '"
            ssql = ssql & Trim(vgCATEGORIA) & "', '"
            ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", "
            ssql = ssql & Trim(vIDCampana) & ", '"
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
            ssql = ssql & Trim(vgCodigoDeProceso) & "', '"
            ssql = ssql & Trim(vLote) & "') "
            cn.Execute ssql
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            ll100 = 0
        End If
        DoEvents
    Loop
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    vlineasTotales = Ll
    Ll = 0
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
        ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    DoEvents
        ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        'cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
        cn1.Execute sSPImportacion & " " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vIDCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If rsCMP("UltimaCorridaError") <> "OK" Then
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            lLote = vLote + 1 'para salir del FOR
        Else
                ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
        End If
        rsCMP.Close
    Next lLote
    '!!!!Esta compania informa altas y modificaciones, no hay que dar de bas=ja a las que no informa!!!!
    'cn1.Execute "TM_BajaDePolizas" & " " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
    cn.Execute "DELETE FROM bandejadeentrada.dbo.importaDatosV2 where idcampana=" & 330 & " and idcia=" & 9999652
Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub



