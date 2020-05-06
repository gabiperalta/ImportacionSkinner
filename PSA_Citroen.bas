Attribute VB_Name = "PSA_Citroen"
Option Explicit
Public Sub ImportarPSACitroen()


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
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos"
    vIDCIA = 10000187
    vIDCampana = 430
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
        vgNROPOLIZA = Mid(sLine, 4, 17)
 '      -------------------------------------------------------
'        vCampo = "NroSecuencial"
'        vPosicion = 17
'        vgNROSECUENCIAL = Mid(sLine, 17, 3)
 '      -------------------------------------------------------
        vCampo = "APELLIDOYNOMBRE"
        vPosicion = 20
        vgAPELLIDOYNOMBRE = Mid(sLine, 21, 49)
 '      -------------------------------------------------------
        vCampo = "DOMICILIO"
        vPosicion = 70
        vgDOMICILIO = Mid(sLine, 70, 55)
 '      -------------------------------------------------------
        vCampo = "LOCALIDAD"
        vPosicion = 125
        vgLOCALIDAD = Mid(sLine, 125, 30)
 '      -------------------------------------------------------
        vCampo = "PROVINCIA"
        vPosicion = 155
        vgPROVINCIA = Mid(sLine, 155, 15)
 '      -------------------------------------------------------
        vCampo = "CODIGOPOSTAL"
        vPosicion = 170
        vgCODIGOPOSTAL = Mid(sLine, 170, 8)
 '      -------------------------------------------------------
        vCampo = "FECHAVIGENCIA"
        vPosicion = 184
        vgFECHAVIGENCIA = Mid(sLine, 184, 2) & "/" & Mid(sLine, 182, 2) & "/" & Mid(sLine, 178, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 192
        vgFECHAVENCIMIENTO = Mid(sLine, 192, 2) & "/" & Mid(sLine, 190, 2) & "/" & Mid(sLine, 186, 4)
 '      -------------------------------------------------------
'        vgFECHAALTAOMNIA = Mid(sLine, 1, 10)
'        vgFECHABAJAOMNIA = Mid(sLine, 1, 10)
'        vgIDAUTO = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "MARCADEVEHICULO"
        vPosicion = 194
        vgMARCADEVEHICULO = Mid(sLine, 194, 30)
        If Trim(vgMARCADEVEHICULO) <> "CITROEN" Then
            MsgBox "Este vehiculo no es CITROEN, revise la base"
            Exit Sub
        End If
 '      -------------------------------------------------------
        vCampo = "MODELO"
        vPosicion = 224
        vgMODELO = Mid(sLine, 224, 20)
 '      -------------------------------------------------------
        vCampo = "COLOR"
        vPosicion = 244
        vgCOLOR = Mid(sLine, 244, 15)
 '      -------------------------------------------------------
        vCampo = "AÑO"
        vPosicion = 259
        vgAno = Mid(sLine, 259, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 263
        vgPATENTE = Mid(sLine, 263, 8)
 '      -------------------------------------------------------
        vCampo = "TipodeServicio"
        vPosicion = 271
        vgTipodeServicio = Mid(sLine, 271, 4)
 '      -------------------------------------------------------
         vCampo = "TIPODEVEHICULO"
        vPosicion = 275
        If Trim(Mid(sLine, 275, 2)) = "A" Then
            vgTIPODEVEHICULO = 1
        ElseIf Trim(Mid(sLine, 275, 2)) = "T" Then
            vgTIPODEVEHICULO = 2
        ElseIf Trim(Mid(sLine, 275, 2)) = "P" Then
            vgTIPODEVEHICULO = 3
        ElseIf Trim(Mid(sLine, 275, 2)) = "M" Then
            vgTIPODEVEHICULO = 5
        Else
            vgTIPODEVEHICULO = 0
        End If
 '      -------------------------------------------------------
'        vgIDTIPODECOBERTURA = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVEHICULO"
        vPosicion = 277
        vgCOBERTURAVEHICULO = Mid(sLine, 277, 2)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVIAJERO"
        vPosicion = 279
        vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
 '      -------------------------------------------------------
        vCampo = "TipodeOperacion"
        vPosicion = 281
        vgTipodeOperacion = Mid(sLine, 281, 1)
'        vgOperacion = Mid(sLine, 1, 10)
'        vgCATEGORIA = Mid(sLine, 1, 10)
'        vgASISTENCIAXENFERMEDAD = Mid(sLine, 1, 10)
'        vgCORRIDA = Mid(sLine, 1, 10)
'        vgFECHACORRIDA = Mid(sLine, 1, 10)
'        vgIdCampana = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "Conductor"
        vPosicion = 282
        vgConductor = Mid(sLine, 282, 50)
 '      -------------------------------------------------------
        vCampo = "CodigoDeProductor"
        vPosicion = 332
        vgCodigoDeProductor = Mid(sLine, 332, 5)
 '      -------------------------------------------------------
        vCampo = "CodigoDeServicioVip"
        vPosicion = 337
        vgCodigoDeServicioVip = Mid(sLine, 337, 1)
 '      -------------------------------------------------------
'        vgTipodeDocumento = Mid(sLine, 1, 10)
'        vgNumeroDeDocumento = Mid(sLine, 1, 10)
'        vgTipodeHogar = Mid(sLine, 1, 10)
'        vgIniciodeAnualidad = Mid(sLine, 1, 10)
'        vgPolizaIniciaAnualidad = Mid(sLine, 1, 10)
'        vgTelefono = Mid(sLine, 1, 10)
'        vgNroMotor = Mid(sLine, 1, 10)
'        vgGama = Mid(sLine, 1, 10)
        
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos ("
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
        ssql = ssql & "IdLote )"
        
        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgidcia) & ", '"
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
        ssql = ssql & Trim(vgidcampana) & ", '"
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
        ssql = ssql & Trim(vLote) & "') "
        cn.Execute ssql
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtProcesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            ll100 = 0
        End If
        DoEvents
    Loop
    ImportadordePolizas.txtProcesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    vlineasTotales = Ll
    Ll = 0
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
        ImportadordePolizas.txtProcesando.BackColor = &HC0C0FF
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vIDCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If rsCMP("UltimaCorridaError") <> "OK" Then
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            lLote = vLote + 1 'para salir del FOR
        Else
                ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
        End If
        rsCMP.Close
    Next lLote
    cn1.Execute "TM_BajaDePolizas" & " " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub


