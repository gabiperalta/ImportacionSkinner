Attribute VB_Name = "Iveco"
Option Explicit
Public Sub ImportarIveco()


Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim vCampo As String
Dim vPosicion As Long

On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos"

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)
    Ll = 1
    Do Until tf.AtEndOfStream
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
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
'        vgNROSECUENCIAL = Mid(sLine, 21, 3)
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
        vPosicion = 178
        vgFECHAVIGENCIA = Mid(sLine, 184, 2) & "/" & Mid(sLine, 182, 2) & "/" & Mid(sLine, 178, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 189
        vgFECHAVENCIMIENTO = Mid(sLine, 195, 2) & "/" & Mid(sLine, 193, 2) & "/" & Mid(sLine, 189, 4)
 '      -------------------------------------------------------
'        vgFECHAALTAOMNIA = Mid(sLine, 1, 10)
'        vgFECHABAJAOMNIA = Mid(sLine, 1, 10)
'        vgIDAUTO = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "MARCADEVEHICULO"
        vPosicion = 200
        vgMARCADEVEHICULO = Mid(sLine, 200, 30)
 '      -------------------------------------------------------
        vCampo = "MODELO"
        vPosicion = 230
        vgMODELO = Mid(sLine, 230, 20)
 '      -------------------------------------------------------
        vCampo = "Gama"
        vPosicion = 250
        vgGama = Mid(sLine, 250, 15)
 '      -------------------------------------------------------
        vCampo = "AÑO"
        vPosicion = 265
        vgAno = Mid(sLine, 265, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 276
        vgPATENTE = Mid(sLine, 276, 8)
 '      -------------------------------------------------------
        vCampo = "TipodeServicio"
        vPosicion = 284
        vgTipodeServicio = Mid(sLine, 284, 4)
 '      -------------------------------------------------------
        vCampo = "TIPODEVEHICULO"
        vPosicion = 288
        If IsNumeric(Mid(sLine, 288, 2)) Then
            vgTIPODEVEHICULO = Mid(sLine, 288, 2)
        Else
            vgTIPODEVEHICULO = 0
        End If
 '      -------------------------------------------------------
'        vgIDTIPODECOBERTURA = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVEHICULO"
        vPosicion = 290
        vgCOBERTURAVEHICULO = Mid(sLine, 290, 2)
 '      -------------------------------------------------------
        vCampo = "COBERTURAVIAJERO"
        vPosicion = 292
        vgCOBERTURAVIAJERO = Mid(sLine, 292, 2)
 '      -------------------------------------------------------
        vCampo = "TipodeOperacion"
        vPosicion = 294
        vgTipodeOperacion = Mid(sLine, 294, 1)
'        vgOperacion = Mid(sLine, 1, 10)
'        vgCATEGORIA = Mid(sLine, 1, 10)
'        vgASISTENCIAXENFERMEDAD = Mid(sLine, 1, 10)
'        vgCORRIDA = Mid(sLine, 1, 10)
'        vgFECHACORRIDA = Mid(sLine, 1, 10)
'        vgIdCampana = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
        vCampo = "Conductor"
        vPosicion = 295
        vgConductor = Mid(sLine, 295, 50)
 '      -------------------------------------------------------
        vCampo = "CodigoDeProductor"
        vPosicion = 345
        vgCodigoDeProductor = Mid(sLine, 345, 5)
 '      -------------------------------------------------------
        vCampo = "CodigoDeServicioVip"
        vPosicion = 350
        vgCodigoDeServicioVip = Mid(sLine, 350, 1)
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
        ssql = ssql & "Gama )"
        
        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgIDCIA) & ", '"
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
        ssql = ssql & Trim(vgIdCampana) & ", '"
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
        ssql = ssql & Trim(vgGama) & "') "
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
    cn1.Execute sSPImportacion
Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub


