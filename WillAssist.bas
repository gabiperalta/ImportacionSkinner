Attribute VB_Name = "WillAssist"
Option Explicit
Public Sub ImportarWillAssist()
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim rsprod As New Recordset
Dim ssql As String
Dim sFile As String
Dim v
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Integer
Dim vCampo As String
Dim vPosicion As Long
Dim regMod As Long

On Error Resume Next
vgidCia = lIdCia ' sale del formulario del importador, al hacer click
vgidCampana = lIdCampana ' sale del formulario del importador, al hacer click

TablaTemporal

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
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

        vCampo = "Numero de poliza"
        vPosicion = 1
        vgNROPOLIZA = Trim(Mid(sLine, 1, 13))
 '      -------------------------------------------------------
'        If IsNumeric(Mid(sLine, 1, 14)) Then
'         vgNROPOLIZA = Mid(sLine, 1, 14)
'         Else
'         vgNROPOLIZA = 0
'        End If
 '      -------------------------------------------------------
'        vCampo = "Numero secuencial"
'        vPosicion = 17
'        vgNUMEROSECUENCIAL = Mid(sLine, 14, 7)
'        vgNUMEROSECUENCIAL = Trim(Right(vgNUMEROSECUENCIAL, 3))
 '      -------------------------------------------------------
        vCampo = "APELLIDO Y NOMBRE"
        vPosicion = 14
        'vgAPELLIDOYNOMBRE = Trim(Mid(sLine, 21, 50))
        vgAPELLIDOYNOMBRE = Trim(Mid(sLine, 14, 50))
 '      -------------------------------------------------------
        vCampo = "DOMICILIO"
        vPosicion = 64
        'vgDOMICILIO = Trim(Mid(sLine, 71, 55))
        vgDOMICILIO = Trim(Mid(sLine, 64, 55))
 '      -------------------------------------------------------
        vCampo = "LOCALIDAD"
        vPosicion = 119
        'vgLOCALIDAD = Trim(Mid(sLine, 126, 30))
        vgLOCALIDAD = Trim(Mid(sLine, 119, 30))
 '      -------------------------------------------------------
        vCampo = "PROVINCIA"
        vPosicion = 149
        'vgPROVINCIA = Trim(Mid(sLine, 156, 15))
        vgPROVINCIA = Trim(Mid(sLine, 149, 15))
 '      -------------------------------------------------------
        vCampo = "CODIGO POSTAL"
        vPosicion = 164
        'vgCODIGOPOSTAL = Trim(Mid(sLine, 171, 8))
        vgCODIGOPOSTAL = Trim(Mid(sLine, 164, 8))
 '      -------------------------------------------------------
        vCampo = "FECHA DESDE"
        vPosicion = 172
        'vgFECHAVIGENCIA = Mid(sLine, 185, 2) & "/" & Mid(sLine, 183, 2) & "/" & Mid(sLine, 179, 4)
        vgFECHAVIGENCIA = Mid(sLine, 178, 2) & "/" & Mid(sLine, 176, 2) & "/" & Mid(sLine, 172, 4)
 '      -------------------------------------------------------
        vCampo = "FECHA HASTA"
        vPosicion = 180
        'vgFECHAVENCIMIENTO = Mid(sLine, 193, 2) & "/" & Mid(sLine, 191, 2) & "/" & Mid(sLine, 187, 4)
        vgFECHAVENCIMIENTO = Mid(sLine, 186, 2) & "/" & Mid(sLine, 184, 2) & "/" & Mid(sLine, 180, 4)
 '      -------------------------------------------------------
        vCampo = "ID PRODUCTO"
        vPosicion = 188
        'v = Trim(Mid(sLine, 195, 4))
        v = Trim(Mid(sLine, 188, 4))
            If Len(v) > 0 Then
            ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & vgidCampana & " and idProductoEnCliente = '" & v & "'"
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
 '      -------------------------------------------------------
        vCampo = "MARCA DE VEHICULO"
        vPosicion = 192
        'vgMARCADEVEHICULO = Trim(Mid(sLine, 199, 30))
        vgMARCADEVEHICULO = Trim(Mid(sLine, 192, 30))
 '      -------------------------------------------------------
        vCampo = "MODELO"
        vPosicion = 222
        'vgMODELO = Trim(Mid(sLine, 229, 20))
        vgMODELO = Trim(Mid(sLine, 222, 20))
 '      -------------------------------------------------------
        vCampo = "AÑO DEL VEHICULO"
        vPosicion = 242
        'vgAno = Trim(Mid(sLine, 249, 5))
        vgAno = Trim(Mid(sLine, 242, 4))
 '      -------------------------------------------------------
        vCampo = "COLOR"
        vPosicion = 247
        'vgCOLOR = Trim(Mid(sLine, 254, 14))
        vgCOLOR = Trim(Mid(sLine, 247, 13))
 '      -------------------------------------------------------
        vCampo = "NUMERO DE PATENTE"
        vPosicion = 261
        'vgPATENTE = Trim(Mid(sLine, 268, 8))
        'vgPATENTE = Trim(Mid(sLine, 260, 8))
        vgPATENTE = Trim(Mid(sLine, 261, 8))
 '      -------------------------------------------------------
        vCampo = "TIPODEVEHICULO"
        vPosicion = 268
        If Trim(Mid(sLine, 268, 2)) = "1" Then
        'If Trim(Mid(sLine, 276, 2)) = "1" Then
            vgTIPODEVEHICULO = 1
        ElseIf Trim(Mid(sLine, 268, 2)) = "2" Then
        'ElseIf Trim(Mid(sLine, 276, 2)) = "2" Then
            vgTIPODEVEHICULO = 2
        ElseIf Trim(Mid(sLine, 268, 2)) = "3" Then
        'ElseIf Trim(Mid(sLine, 276, 2)) = "3" Then
            vgTIPODEVEHICULO = 3
        ElseIf Trim(Mid(sLine, 268, 2)) = "4" Then
        'ElseIf Trim(Mid(sLine, 276, 2)) = "4" Then
            vgTIPODEVEHICULO = 4
        ElseIf Trim(Mid(sLine, 268, 2)) = "5" Then
        'ElseIf Trim(Mid(sLine, 276, 2)) = "5" Then
            vgTIPODEVEHICULO = 5
        Else
            vgTIPODEVEHICULO = 0
        End If


 '      -------------------------------------------------------
        vCampo = "Conductor"
        vPosicion = 271
        vgConductor = Trim(Mid(sLine, 270, 50))
        'vgConductor = Trim(Mid(sLine, 278, 50))
 '      -------------------------------------------------------
        vCampo = "Agencia - Sponsor"
        vPosicion = 321
        vgAgencia = Trim(Mid(sLine, 320, 10))
        'vgAgencia = Trim(Mid(sLine, 328, 10))
 '      -------------------------------------------------------
        vCampo = "CodigoDeServicioVip"
        vPosicion = 331
        vgCodigoDeServicioVip = Trim(Mid(sLine, 330, 1))
        'vgCodigoDeServicioVip = Trim(Mid(sLine, 338, 1))
 '      -------------------------------------------------------
        vCampo = "EMAIL"
        vPosicion = 332
        vgEmail = Trim(Mid(sLine, 331, 80)) '50
        'vgEmail = Trim(Mid(sLine, 339, 80))
 '      -------------------------------------------------------

    Dim rscn1 As New Recordset
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' and PATENTE = '" & Trim(vgPATENTE) & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
'                If Trim(rscn1("NROSECUENCIAL")) <> Trim(vgNUMEROSECUENCIAL) Then vdif = vdif + 1
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
'                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
                If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                If Trim(rscn1("Email")) <> Trim(vgEmail) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("CodigoEnCliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
                If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
                If Trim(rscn1("Agencia")) <> Trim(vgAgencia) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If
        rscn1.Close
'-=================================================================================================================
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
        ssql = ssql & "IDPOLIZA, "
        ssql = ssql & "IDCIA, "
        ssql = ssql & "NROPOLIZA, "
        ssql = ssql & "NROSECUENCIAL, "
        ssql = ssql & "APELLIDOYNOMBRE, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "CODIGOPOSTAL, "
        ssql = ssql & "Email, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "FECHAVENCIMIENTO, "
        ssql = ssql & "MARCADEVEHICULO, "
        ssql = ssql & "MODELO, "
        ssql = ssql & "COBERTURAVEHICULO, "
        ssql = ssql & "COBERTURAVIAJERO, "
        ssql = ssql & "COBERTURAHOGAR, "
        ssql = ssql & "COLOR, "
        ssql = ssql & "ANO, "
        ssql = ssql & "PATENTE, "
        ssql = ssql & "TIPODEVEHICULO, "
        ssql = ssql & "CORRIDA, "
        ssql = ssql & "IdCampana, "
        ssql = ssql & "Conductor, "
        ssql = ssql & "Agencia, "
        ssql = ssql & "CodigoDeServicioVip, "
        ssql = ssql & "CodigoEnCliente, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "Modificaciones)"
        
        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgidCia) & ", '"
        ssql = ssql & Trim(vgNROPOLIZA) & "', '"
        ssql = ssql & Trim(vgNUMEROSECUENCIAL) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgLOCALIDAD) & "', '"
        ssql = ssql & Trim(vgPROVINCIA) & "', '"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
        ssql = ssql & Trim(vgEmail) & "', '"
        ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
        ssql = ssql & Trim(vgMODELO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
        ssql = ssql & Trim(vgCOLOR) & "', '"
        ssql = ssql & Trim(vgAno) & "', '"
        ssql = ssql & Trim(vgPATENTE) & "', "
        ssql = ssql & Trim(vgTIPODEVEHICULO) & ", "
        ssql = ssql & Trim(vgCORRIDA) & ", "
        ssql = ssql & Trim(vgidCampana) & ", '"
        ssql = ssql & Trim(vgConductor) & "', '"
        ssql = ssql & Trim(vgAgencia) & "', '"
        ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
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
                
                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lIdCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
                
            ll100 = 0
        End If
        DoEvents
    Loop
    
'================Control de Leidos===========llama al storeprocedure para hacer un update en tm_importacionHistorial

    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
    listoParaProcesar
    
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
'===============inicio del Control de Procesos=====================================

    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    Dim rsCMP As New Recordset
    DoEvents
    For lLote = 1 To vLote
            cn1.CommandTimeout = 300
            cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & lIdCia & ", " & lIdCampana ' & ", " & vNombreTablaTemporal
            ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & lIdCampana
            rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " Procesando los datos"
            ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
            DoEvents
            rsCMP.Close
        Next lLote
   
    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & vgidCampana & ", " & vgCORRIDA
    Procesado
'====================================================================================
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



