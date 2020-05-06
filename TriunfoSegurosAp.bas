Attribute VB_Name = "TriunfoSegurosAP"
Option Explicit
Public Sub ImportarTriunfoSegurosAp()


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
Dim vSinImportar As String
Dim vCoberturaFinanciera As String
Dim vFechaDeNacimiento As String
Dim vFechaDeVigencia As String
Dim vFechaDeVencimiento As String
'Dim rs As ADODB.Recordset
'    rs.Open Null, Null, adOpenKeyset, adLockBatchOptimistic
'    rs.batchupdate

On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosTriunfoSegurosAP"
    vIDCIA = 509
    vIDCampana = 726
    LongDeLote = 1000

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)
    Ll = 1
    nroLinea = 1
    vLote = 1
    tf.SkipLine
    Do Until tf.AtEndOfStream
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
        
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
        
'Tipdoc
        vCampo = "Tipdoc"
        vPosicion = 1
        vgTipodeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'numdoc
        vCampo = "numdoc"
        vPosicion = 2
        vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Nombre Beneficiario
        vCampo = "Nombre Beneficiario"
        vPosicion = 3
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'AñoNac
        vCampo = "AñoNac"
        vPosicion = 4
        vFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'MesNac
        vCampo = "MesNac"
        vPosicion = 5
        vFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeNacimiento
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'DiaNac
        vCampo = "DiaNac"
        vPosicion = 6
        vFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeNacimiento
        If IsDate(vFechaDeNacimiento) Then vgFechaDeNacimiento = vFechaDeNacimiento
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Rama
        vCampo = "Rama"
        vPosicion = 7
        vgTipodeOperacion = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Certificado
        vCampo = "Certificado"
        vPosicion = 8
        vgCertificado = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Poliza
        vCampo = "Poliza"
        vPosicion = 9
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        vgNROPOLIZA = vgTipodeOperacion & "-" & vgCertificado & "-" & vgNROPOLIZA
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Año Vig Des
        vCampo = "Año Vig Des"
        vPosicion = 10
        vFechaDeVigencia = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Mes Vig Des
        vCampo = "Mes Vig Des"
        vPosicion = 11
        vFechaDeVigencia = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeVigencia
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dia Vig des
        vCampo = "Dia Vig des"
        vPosicion = 12
        vFechaDeVigencia = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeVigencia
        If IsDate(vFechaDeVigencia) Then vgFECHAVIGENCIA = vFechaDeVigencia
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Año Vig Has
        vCampo = "Año Vig Has"
        vPosicion = 13
        vFechaDeVencimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Mes Vig Has
        vCampo = "Mes Vig Has"
        vPosicion = 14
        vFechaDeVencimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeVencimiento
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dia Vig Has
        vCampo = "Dia Vig Has"
        vPosicion = 15
        vFechaDeVencimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1) & "/" & vFechaDeVencimiento
        If IsDate(vFechaDeVencimiento) Then vgFECHAVENCIMIENTO = vFechaDeVencimiento
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Año Alta Nomi
        vCampo = "Año Alta Nomi"
        vPosicion = 16
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Mes Alta Nomi
        vCampo = "Mes Alta Nomi"
        vPosicion = 17
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dia Alta Nomi
        vCampo = "Dia Alta Nomi"
        vPosicion = 18
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Año Baja Nomi
        vCampo = "Año Baja Nomi"
        vPosicion = 19
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Mes Baja Nomi
        vCampo = "Mes Baja Nomi"
        vPosicion = 20
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dia Baja Nomi
        vCampo = "Dia Baja Nomi"
        vPosicion = 21
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Domicilio Beneficiario2
        vCampo = "Domicilio Beneficiario2"
        vPosicion = 22
        vgDOMICILIO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Numero
        vCampo = "Numero"
        vPosicion = 23
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Piso
        vCampo = "Piso"
        vPosicion = 24
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dpto
        vCampo = "Dpto"
        vPosicion = 25
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Cod Pos
        vCampo = "Cod Pos"
        vPosicion = 26
        vgCODIGOPOSTAL = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Localidad
        vCampo = "Localidad"
        vPosicion = 27
        vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Provincia
        vCampo = "Provincia"
        vPosicion = 28
        vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Ocupacion
        vCampo = "Ocupacion"
        vPosicion = 29
        vgOcupacion = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Suma Aseg.
        vCampo = "Suma Aseg."
        vPosicion = 30
        If IsNumeric(Mid(sLine, InStr(1, sLine, ";") + 1)) Then
            vgImporte = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        Else
            vgImporte = 0
        End If
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Nombre Tomador
        vCampo = "Nombre Tomador"
        vPosicion = 31
        vSinImportar = Mid(sLine, 1, Len(sLine))
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Año Novedad
        vCampo = "Año Novedad"
        vPosicion = 32
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Mes Novedad
        vCampo = "Mes Novedad"
        vPosicion = 33
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dia Novedad
        vCampo = "Dia Novedad"
        vPosicion = 34
        vSinImportar = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Origen
        vCampo = "Origen"
        vPosicion = 35
        vgorigen = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Cob. Cobranza
        vCampo = "Cob. Cobranza"
        vPosicion = 36
        vCoberturaFinanciera = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Columna1
        vCampo = "Columna1"
        vPosicion = 37
        vgCodigoEnCliente = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Columna2
        vCampo = "Columna2"
        vPosicion = 38
        vgNroSecunecialEnCliente = Mid(sLine, 1)
        If InStr(1, sLine, ";") > 0 Then
            MsgBox "Campos con ; que no corresponden en la linea" & Ll
            Exit Sub
        End If

        
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatosTriunfoSegurosAP ("
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
        ssql = ssql & "FechaDeNacimiento, "
        ssql = ssql & "InformadoSinCobertura, "
        ssql = ssql & "Ocupacion, "
        ssql = ssql & "Importe, "
        ssql = ssql & "Certificado,CodigoEnCliente,NroSecunecialEnCliente )"

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
        ssql = ssql & Trim(vgANO) & "', '"
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
        ssql = ssql & Trim(vgGama) & "', '"
        ssql = ssql & Trim(vLote) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vgOcupacion) & "', "
        ssql = ssql & Trim(vgImporte) & ", '"
        ssql = ssql & Trim(vCoberturaFinanciera) & ", '"
        ssql = ssql & Trim(vgCertificado) & "', '"
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
        ssql = ssql & Trim(vgNroSecunecialEnCliente) & "') "

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
    If MsgBox("¿Desea Procesar los datos ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    vlineasTotales = Ll
    Ll = 0
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vIDCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
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



