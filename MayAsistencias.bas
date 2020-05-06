Attribute VB_Name = "MayAsistencias"
Option Explicit
Public Sub ImportarMayAsistencias()


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
'Dim rs As ADODB.Recordset
'    rs.Open Null, Null, adOpenKeyset, adLockBatchOptimistic
'    rs.batchupdate
Dim vpVig As String
Dim vpVen As String
Dim vpTIPODEVEHICULO As String
On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosMayAsistencias"
    vIDCIA = 10000572
    vIDCampana = 525
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
        vgPATENTE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNroMotor = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgAPELLIDOYNOMBRE = vgAPELLIDOYNOMBRE + ", " + Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgTipodeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
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
        vgCOLOR = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgCOBERTURAVEHICULO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vpVig = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        If Not IsDate(Mid(vpVig, 5, 2) & "/" & Mid(vpVig, 3, 2) & "/20" & Mid(vpVig, 1, 2)) Then
            MsgBox "Fecha de inicio invalida en Linea " & Ll
            Exit Sub
        End If
        vgFECHAVIGENCIA = Mid(vpVig, 5, 2) & "/" & Mid(vpVig, 3, 2) & "/20" & Mid(vpVig, 1, 2)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vpVen = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        If Not IsDate(Mid(vpVen, 5, 2) & "/" & Mid(vpVen, 3, 2) & "/20" & Mid(vpVen, 1, 2)) Then
            MsgBox "Fecha de final invalida en Linea " & Ll
            Exit Sub
        End If
        vgFECHAVENCIMIENTO = Mid(vpVen, 5, 2) & "/" & Mid(vpVen, 3, 2) & "/20" & Mid(vpVen, 1, 2)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vgOperacion = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
            sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
        vpTIPODEVEHICULO = sLine
        Select Case vpTIPODEVEHICULO
            Case "LIVIANO"
                vgTIPODEVEHICULO = "1"
            Case Else
                vgTIPODEVEHICULO = "4"
        
        End Select

'        Select Case vgPROVINCIA
'            Case "01"
'            vgPROVINCIA = "CAPITAL"
'            Case "02"
'            vgPROVINCIA = "BUENOS AIRES"
'            Case "03"
'            vgPROVINCIA = "CATAMARCA"
'            Case "04"
'            vgPROVINCIA = "CORDOBA"
'            Case "05"
'            vgPROVINCIA = "CORRIENTES"
'            Case "06"
'            vgPROVINCIA = "ENTRE RIOS"
'            Case "07"
'            vgPROVINCIA = "JUJUY"
'            Case "08"
'            vgPROVINCIA = "LA RIOJA"
'            Case "09"
'            vgPROVINCIA = "MENDOZA"
'            Case "10"
'            vgPROVINCIA = "SALTA"
'            Case "11"
'            vgPROVINCIA = "SAN JUAN"
'            Case "12"
'            vgPROVINCIA = "SAN LUIS"
'            Case "13"
'            vgPROVINCIA = "SANTA FE"
'            Case "14"
'            vgPROVINCIA = "SANTIAGO DEL ESTERO"
'            Case "15"
'            vgPROVINCIA = "TUCUMAN"
'            Case "16"
'            vgPROVINCIA = "CHACO"
'            Case "17"
'            vgPROVINCIA = "CHUBUT"
'            Case "18"
'            vgPROVINCIA = "FORMOSA"
'            Case "19"
'            vgPROVINCIA = "LA PAMPA"
'            Case "20"
'            vgPROVINCIA = "MISIONES"
'            Case "21"
'            vgPROVINCIA = "MISIONES"
'            Case "22"
'            vgPROVINCIA = "Neuquen"
'            Case "23"
'            vgPROVINCIA = "Rio Negro"
'            Case "24"
'            vgPROVINCIA = "Santa Cruz"
'            Case "25"
'            vgPROVINCIA = "TierraDelFuego"
'            Case "26"
'            vgPROVINCIA = "Exterior"
'        End Select
 '      -------------------------------------------------------

'        If DateDiff("d", vgFECHAVENCIMIENTO, Now()) < 0 Then
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatosMayAsistencias ("
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
'        End If
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
        ImportadordePolizas.txtProcesando.BackColor = &HC0C0FF
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




