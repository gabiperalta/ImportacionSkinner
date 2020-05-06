Attribute VB_Name = "NUEVOBANCOCOMERCIALDEURUGUAY"
Option Explicit
Public Sub ImportarNUEVOBANCOCOMERCIALDEURUGUAY()


Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim vCampo As String
Dim vPosicion As Long
Dim vnsec As String

On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosNBCU"

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
'        vCampo = "Compania"
'        vPosicion = 1
'        vgNUMEROCOMPANIA = Mid(sLine, 1, 3)
 '      -------------------------------------------------------
        vCampo = "NroPoliza"
        vPosicion = 10
        vgNROPOLIZA = Mid(sLine, 10, 4) & "_" & Mid(sLine, 22, 4) & "-" & Mid(sLine, 26, 8) & Mid(sLine, 92, 2)
 '      -------------------------------------------------------
        vCampo = "NumeroDeDocumento"
        vPosicion = 26
        vgNumeroDeDocumento = Mid(sLine, 26, 8)
 '      -------------------------------------------------------
        vCampo = "APELLIDOYNOMBRE"
        vPosicion = 109
        vgAPELLIDOYNOMBRE = Mid(sLine, 109, 60)
 '      -------------------------------------------------------
'        vCampo = "DOMICILIO"
'        vPosicion = 70
'        vgDOMICILIO = Mid(sLine, 70, 55)
 '      -------------------------------------------------------
        vCampo = "LOCALIDAD"
        vPosicion = 46
        vgLOCALIDAD = Mid(sLine, 46, 35)
 '      -------------------------------------------------------
'        vCampo = "PROVINCIA"
'        vPosicion = 155
'        vgPROVINCIA = Mid(sLine, 155, 15)
 '      -------------------------------------------------------
'        vCampo = "CODIGOPOSTAL"
'        vPosicion = 170
'        vgCODIGOPOSTAL = Mid(sLine, 170, 8)
 '      -------------------------------------------------------
        vCampo = "FECHAVIGENCIA"
        vPosicion = 82
        vgFECHAVIGENCIA = Mid(sLine, 88, 2) & "/" & Mid(sLine, 86, 2) & "/" & Mid(sLine, 82, 4)
 '      -------------------------------------------------------
        vCampo = "FECHAVENCIMIENTO"
        vPosicion = 90
        vgFECHAVENCIMIENTO = Mid(sLine, 96, 2) & "/" & Mid(sLine, 94, 2) & "/" & Mid(sLine, 90, 4)
 '      -------------------------------------------------------
        If Len(Trim(Mid(sLine, 101, 4))) > 0 Then
            vCampo = "FECHABAJAOMNIA"
            vPosicion = 101
            vgFECHABAJAOMNIA = Mid(sLine, 107, 2) & "/" & Mid(sLine, 105, 2) & "/" & Mid(sLine, 101, 4)
        Else
            vgFECHABAJAOMNIA = "01/01/2100"
        End If
 '      -------------------------------------------------------
''        vgFECHAALTAOMNIA = Mid(sLine, 1, 10)
''        vgFECHABAJAOMNIA = Mid(sLine, 1, 10)
''        vgIDAUTO = Mid(sLine, 1, 10)
' '      -------------------------------------------------------
'        vCampo = "MARCADEVEHICULO"
'        vPosicion = 194
'        vgMARCADEVEHICULO = Mid(sLine, 194, 30)
' '      -------------------------------------------------------
'        vCampo = "MODELO"
'        vPosicion = 224
'        vgMODELO = Mid(sLine, 224, 20)
' '      -------------------------------------------------------
'        vCampo = "COLOR"
'        vPosicion = 244
'        vgCOLOR = Mid(sLine, 244, 15)
' '      -------------------------------------------------------
'        vCampo = "AÑO"
'        vPosicion = 259
'        vgANO = Mid(sLine, 259, 4)
 '      -------------------------------------------------------
        vCampo = "PATENTE"
        vPosicion = 1
        vgPATENTE = Mid(sLine, 1, 9)
 '      -------------------------------------------------------
        vCampo = "NroSecuencial"
        vPosicion = 17
        If Mid(sLine, 81, 1) = "T" Then
            vnsec = "1"
        Else
            vnsec = "0"
        End If
        vgNROSECUENCIAL = Mid(sLine, 94, 2) & vnsec
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
'        vCampo = "COBERTURAVEHICULO"
'        vPosicion = 277
'        vgCOBERTURAVEHICULO = Mid(sLine, 277, 2)
' '      -------------------------------------------------------
'        vCampo = "COBERTURAVIAJERO"
'        vPosicion = 279
'        vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
' '      -------------------------------------------------------
        vCampo = "Operacion"
        vPosicion = 98
        vgOperacion = Mid(sLine, 98, 3)
'        vgTipodeOperacion = Mid(sLine, 281, 1)
'        vgCATEGORIA = Mid(sLine, 1, 10)
'        vgASISTENCIAXENFERMEDAD = Mid(sLine, 1, 10)
'        vgCORRIDA = Mid(sLine, 1, 10)
'        vgFECHACORRIDA = Mid(sLine, 1, 10)
'        vgIdCampana = Mid(sLine, 1, 10)
 '      -------------------------------------------------------
'        vCampo = "Conductor"
'        vPosicion = 282
'        vgConductor = Mid(sLine, 282, 50)
' '      -------------------------------------------------------
'        vCampo = "CodigoDeProductor"
'        vPosicion = 332
'        vgCodigoDeProductor = Mid(sLine, 332, 5)
' '      -------------------------------------------------------
'        vCampo = "CodigoDeServicioVip"
'        vPosicion = 337
'        vgCodigoDeServicioVip = Mid(sLine, 337, 1)
 '      -------------------------------------------------------
'        vgTipodeDocumento = Mid(sLine, 1, 10)
'        vgNumeroDeDocumento = Mid(sLine, 1, 10)
'        vgTipodeHogar = Mid(sLine, 1, 10)
'        vgIniciodeAnualidad = Mid(sLine, 1, 10)
'        vgPolizaIniciaAnualidad = Mid(sLine, 1, 10)
'        vgTelefono = Mid(sLine, 1, 10)
'        vgNroMotor = Mid(sLine, 1, 10)
'        vgGama = Mid(sLine, 1, 10)
        
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatosNBCU ("
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
        ssql = ssql & "FECHABAJAOMNIA, "
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
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgFECHABAJAOMNIA) & "', "
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



