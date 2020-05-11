Attribute VB_Name = "Segurcoop"
Option Explicit
Public Sub ImportarSegurcoop()
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
Dim sssql As String
Dim rsprod As New Recordset
Dim lRow
Dim sName
Dim sFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long, ll100 As Long
'Dim nroLinea As Long
Dim vCampo As String
Dim vPosicion As Long
'Dim lLote As Long
'Dim vLote As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
'Dim vIDCIA As Long
'Dim vIDCampana As Long
'Dim rsCMP As New Recordset
'Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim vLTIPODEVEHICULO As String
Dim vTipodeServicio As String
Dim vTipodeServicioActual As String
Dim regMod As Long
Dim nroArchivo As Integer
Dim archivoLeido As Integer

Dim vLeidosLista(0 To 3) As Long
'Dim vCoberturaLista(0 To 3, 0 To 2, 0 To 20) As String            ' campana,tipoCobertura,coberturaEncontrada
'Dim vLeidosPorCoberturaLista(0 To 3, 0 To 2, 0 To 20) As Long     ' campana,tipoCobertura,coberturaEncontrada
Dim vLeidosPorCoberturaLista(0 To 2) As Long                       ' tipoCobertura
Dim vLotesLista(0 To 3) As Long
Dim vidCampanaLista(0 To 3) As Integer
Dim vidHistorialImportacionLista(0 To 3) As Long
Dim finProcesoCampana As Integer
Dim vCodigoEnClienteActual As String

Dim sAno As Date
Dim sMes As Date
Dim sDia As Date
Dim sZone As String
Dim sArchivo As String
Dim sExtencion As String

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

vidCampanaLista(0) = 1073   'ATM
vidCampanaLista(1) = 1075   'Viajero
vidCampanaLista(2) = 1076   'Hogar/Comercios
vidCampanaLista(3) = 1077   'Vehiculos/Camiones

vCodigoEnClienteActual = "ATM"

'=====Creacion de la tabla temporal================
TablaTemporal 'vgidCampana = 1073
'==================================================

''=====Prueba de optimizacion de tablas=============
'TablaTemporalOptimizadaPrueba
''vgidCampana = 1075
''TablaTemporal
''vgidCampana = 1076
''TablaTemporalOptimizadaPrueba
'
''vgidCampana = 1073
''==================================================

Dim vCantDeErrores As Long
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores" 'creacion de archivo .log que almacena los errores surgidos de la importacion ( registros no importados)
vCantDeErrores = 0

If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

Ll = 0
ll100 = 0
regMod = 0
finProcesoCampana = 1 'por ATM

nroArchivo = 0
archivoLeido = 0
sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
sArchivo = Mid(fileimportacion, 1, Len(fileimportacion) - 4)
sExtencion = Mid(fileimportacion, InStr(1, fileimportacion, "."), Len(fileimportacion) + InStr(1, fileimportacion, "."))

'Dim nroArchivoNuevo As Integer
For nroArchivo = 0 To 8

    If nroArchivo = 0 Then
    
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
            vidHistorialImportacionLista(0) = vgIdHistorialImportacion
            rsCorr.Close
            
        End If
        '=======seteo control de lote================================================================
        Dim lLote As Long
        Dim vLote As Long
        Dim nroLinea As Long
        Dim LongDeLote As Long
        Dim totalRegistros As Long
        LongDeLote = 1000
        nroLinea = 1
        vLote = 1
        totalRegistros = 1
        '============================================================================================
        Ll = 1
        nroLinea = 1
        vLote = 1

    End If

    Select Case nroArchivo
        Case 0
            vCodigoEnClienteActual = "ATM"
        Case 1
            Ll = 1
            nroLinea = 1
            vLote = 1
            ll100 = 0
            regMod = 0
            fileimportacion = "viaj.txt"
            vgidCampana = vidCampanaLista(1)
            vCodigoEnClienteActual = "EA"
            finProcesoCampana = 1
        Case 2
            finProcesoCampana = 0
            Ll = 1
            nroLinea = 1
            vLote = 1
            ll100 = 0
            regMod = 0
            fileimportacion = "COMBINADO.txt"
            vgidCampana = vidCampanaLista(2)
            vCodigoEnClienteActual = "HOG"
        Case 3
            fileimportacion = "INTEGRAL.txt"
            vgidCampana = vidCampanaLista(2)
            vCodigoEnClienteActual = "INT"
            finProcesoCampana = 1
        Case 4
            finProcesoCampana = 0
            Ll = 1
            nroLinea = 1
            vLote = 1
            ll100 = 0
            regMod = 0
            fileimportacion = "AUTO_A.txt"
            vgidCampana = vidCampanaLista(3)
            vCodigoEnClienteActual = "AUA"
        Case 5
            fileimportacion = "AUTO_B.txt"
            vgidCampana = vidCampanaLista(3)
            vCodigoEnClienteActual = "AUB"
        Case 6
            fileimportacion = "AUTO_C.txt"
            vgidCampana = vidCampanaLista(3)
            vCodigoEnClienteActual = "AUC"
        Case 7
            fileimportacion = "AUTO_E.txt"
            vgidCampana = vidCampanaLista(3)
            vCodigoEnClienteActual = "AUE"
        Case 8
            fileimportacion = "CAMIONES.txt"
            vgidCampana = vidCampanaLista(3)
            vCodigoEnClienteActual = "CA"
            finProcesoCampana = 1
        Case Else
            Exit For
    End Select
    
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion

    If fs.FileExists(sFile) Then
        
        archivoLeido = 1
        
        Set tf = fs.OpenTextFile(sFile, ForReading, True)
        
        '=======control de lectura del archivo de datos=======================
        If Err Then
            MsgBox Err.Description
            Err.Clear
            Exit Sub
        End If
        
        If nroArchivo = 1 Or nroArchivo = 2 Or nroArchivo = 4 Then
            cn1.Execute "TM_CargaPolizasLogDeSetCorridasSegurcoop " & vgidCampana & ", " & vgCORRIDA
            
            ' se obtiene el idHistorialDeImportacion
            ssql = "SELECT IdHistorialDeImportacion "
            ssql = ssql & " FROM tm_ImportacionHistorial "
            ssql = ssql & " WHERE corrida = " & vgCORRIDA & " and idCampana = " & vgidCampana
            
            rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vgIdHistorialImportacion = rsCorr("IdHistorialDeImportacion")
            rsCorr.Close
            
            Select Case nroArchivo
                Case 1
                    vidHistorialImportacionLista(1) = vgIdHistorialImportacion
                Case 2
                    vidHistorialImportacionLista(2) = vgIdHistorialImportacion
                Case 4
                    vidHistorialImportacionLista(3) = vgIdHistorialImportacion
            End Select
            
        End If
        
        Do Until tf.AtEndOfStream
            
            Blanquear
            
            vgCodigoEnCliente = vCodigoEnClienteActual
            
            sLine = tf.ReadLine
            If Len(Trim(sLine)) < 5 Then Exit Do
            sLine = Replace(sLine, "'", "*")
            
            '=======Control de Lote===============================
            nroLinea = nroLinea + 1
            If nroLinea = LongDeLote + 1 Then
                vLote = vLote + 1
                nroLinea = 1
            End If
            
            totalRegistros = totalRegistros + 1
            
            '=======Se revisa que archivo se esta procesando======
            '   ATM
            If vgidCampana = 1073 Then
                vCampo = "NRO. EMPRESA"
                vPosicion = 1
                vgNUMEROCOMPANIA = Mid(sLine, vPosicion, 3)
                '      -------------------------------------------------------
                vCampo = "NRO. POLIZA"
                vPosicion = 4
                vgNROPOLIZA = Val(Trim(Mid(sLine, vPosicion, 20)))
                '      -------------------------------------------------------
                vCampo = "ITEM"
                vPosicion = 24
                vgNROSECUENCIAL = Mid(sLine, vPosicion, 6)
                vgNROPOLIZA = Trim(vgNROPOLIZA) & Trim(vgNROSECUENCIAL)
                '      -------------------------------------------------------
                vCampo = "APELLIDOYNOMBRE"
                vPosicion = 30
                vgAPELLIDOYNOMBRE = Mid(sLine, vPosicion, 50)
                '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 80
                vgDOMICILIO = Mid(sLine, vPosicion, 55)
                '      -------------------------------------------------------
                vCampo = "COD. POSTAL"
                vPosicion = 135
                vgCODIGOPOSTAL = Mid(sLine, vPosicion, 8)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA DESDE"
                vPosicion = 143
                vgFECHAVIGENCIA = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA HASTA"
                vPosicion = 151
                vgFECHAVENCIMIENTO = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                'vCampo = "MARCA VEHICULO"
                'vPosicion = 184
                'vgFECHAVIGENCIA = Mid(sLine, 184, 2) & "/" & Mid(sLine, 182, 2) & "/" & Mid(sLine, 178, 4)
                '      -------------------------------------------------------
                'vCampo = "MODELO"
                'vPosicion = 192
                'vgFECHAVENCIMIENTO = Mid(sLine, 192, 2) & "/" & Mid(sLine, 190, 2) & "/" & Mid(sLine, 186, 4)
                '      -------------------------------------------------------
                'vCampo = "COLOR"
                'vPosicion = 194
                'vgMARCADEVEHICULO = Mid(sLine, 194, 30)
                '      -------------------------------------------------------
                vCampo = "AÑO"
                vPosicion = 239
                vgTipodeDocumento = Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "PATENTE"
                vPosicion = 243
                vgNumeroDeDocumento = Val(Trim(Mid(sLine, vPosicion, 15)))
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 263
                'vgPATENTE = Mid(sLine, 263, 8)
                '      -------------------------------------------------------
                'vCampo = "COB. AUTO"
                'vPosicion = 277
                'vgCOBERTURAVEHICULO = Mid(sLine, 277, 2)
                '      -------------------------------------------------------
                'vCampo = "COB. VIAJ."
                'vPosicion = 279
                'vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
                '      -------------------------------------------------------
                'vCampo = "COB. HOGAR"
                'vPosicion = 279
                'vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
                '      -------------------------------------------------------
                vCampo = "SIN USO"
                vPosicion = 263
                'vgCodigoEnCliente = Trim(Mid(sLine, vPosicion, 10))
                '      -------------------------------------------------------
                'vCampo = "PREST.ADIC."
                'vPosicion = 282
                'vgConductor = Mid(sLine, 282, 50)
                '      -------------------------------------------------------
                'vCampo = "COD.PROC."
                'vPosicion = 332
                'vgCodigoDeProductor = Mid(sLine, 332, 5)
            
            '   Combinado integral
            ElseIf vgidCampana = 1076 Then
                vCampo = "NRO. EMPRESA"
                vPosicion = 1
                vgNUMEROCOMPANIA = Mid(sLine, vPosicion, 3)
                '      -------------------------------------------------------
                vCampo = "NRO. POLIZA"
                vPosicion = 4
                vgNROPOLIZA = Val(Trim(Mid(sLine, vPosicion, 20)))
                '      -------------------------------------------------------
                vCampo = "ITEM"
                vPosicion = 24
                vgNROSECUENCIAL = Mid(sLine, vPosicion, 6)
                vgNROPOLIZA = Trim(vgNROPOLIZA) & Trim(vgNROSECUENCIAL)
                '      -------------------------------------------------------
                vCampo = "APELLIDOYNOMBRE"
                vPosicion = 30
                vgAPELLIDOYNOMBRE = Mid(sLine, vPosicion, 50)
                '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 80
                vgDOMICILIO = Mid(sLine, vPosicion, 55)
                '      -------------------------------------------------------
                vCampo = "COD. POSTAL"
                vPosicion = 135
                vgCODIGOPOSTAL = Mid(sLine, vPosicion, 8)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA DESDE"
                vPosicion = 143
                vgFECHAVIGENCIA = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA HASTA"
                vPosicion = 151
                vgFECHAVENCIMIENTO = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                'vCampo = "MARCA VEHICULO"
                'vPosicion = 184
                'vgFECHAVIGENCIA = Mid(sLine, 184, 2) & "/" & Mid(sLine, 182, 2) & "/" & Mid(sLine, 178, 4)
                '      -------------------------------------------------------
                'vCampo = "MODELO"
                'vPosicion = 192
                'vgFECHAVENCIMIENTO = Mid(sLine, 192, 2) & "/" & Mid(sLine, 190, 2) & "/" & Mid(sLine, 186, 4)
                '      -------------------------------------------------------
                'vCampo = "COLOR"
                'vPosicion = 194
                'vgMARCADEVEHICULO = Mid(sLine, 194, 30)
                '      -------------------------------------------------------
                'vCampo = "AÑO"
                'vPosicion = 239
                'vgTipodeDocumento = Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                'vCampo = "PATENTE"
                'vPosicion = 243
                'vgNumeroDeDocumento = Mid(sLine, vPosicion, 15)
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 263
                'vgPATENTE = Mid(sLine, 263, 8)
                '      -------------------------------------------------------
                'vCampo = "COB. AUTO"
                'vPosicion = 277
                'vgCOBERTURAVEHICULO = Mid(sLine, 277, 2)
                '      -------------------------------------------------------
                'vCampo = "COB. VIAJ."
                'vPosicion = 279
                'vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
                '      -------------------------------------------------------
                vCampo = "COB. HOGAR"
                vPosicion = 259
                'vgCOBERTURAHOGAR = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 263
                '      -------------------------------------------------------
                'vCampo = "PREST.ADIC."
                'vPosicion = 282
                'vgConductor = Mid(sLine, 282, 50)
                '      -------------------------------------------------------
                'vCampo = "COD.PROC."
                'vPosicion = 332
                'vgCodigoDeProductor = Mid(sLine, 332, 5)
                '      -------------------------------------------------------
                vCampo = "CRISTALES"
                vPosicion = 276
                vgTopeCritales = Mid(sLine, vPosicion, 19) ' 16 enteros + coma + 2 decimales
                vgTopeCritales = Replace(vgTopeCritales, ".", ",") ' en subdesarrollo, el double es con "," y no con "."
                vgTopeCritales = CStr(CDbl(vgTopeCritales))
                vgTopeCritales = Replace(vgTopeCritales, ",", ".") ' en subdesarrollo, el double es con "," y no con "."
    
            '   Autos y camiones
            ElseIf vgidCampana = 1077 Then
                vCampo = "NRO. EMPRESA"
                vPosicion = 1
                vgNUMEROCOMPANIA = Mid(sLine, vPosicion, 3)
                '      -------------------------------------------------------
                vCampo = "NRO. POLIZA"
                vPosicion = 4
                vgNROPOLIZA = Val(Trim(Mid(sLine, vPosicion, 20)))
                '      -------------------------------------------------------
                vCampo = "ITEM"
                vPosicion = 24
                vgNROSECUENCIAL = Mid(sLine, vPosicion, 6)
                vgNROPOLIZA = Trim(vgNROPOLIZA) & Trim(vgNROSECUENCIAL)
                '      -------------------------------------------------------
                vCampo = "APELLIDOYNOMBRE"
                vPosicion = 30
                vgAPELLIDOYNOMBRE = Mid(sLine, vPosicion, 50)
                '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 80
                vgDOMICILIO = Mid(sLine, vPosicion, 55)
                '      -------------------------------------------------------
                vCampo = "COD. POSTAL"
                vPosicion = 135
                vgCODIGOPOSTAL = Mid(sLine, vPosicion, 8)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA DESDE"
                vPosicion = 143
                vgFECHAVIGENCIA = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA HASTA"
                vPosicion = 151
                vgFECHAVENCIMIENTO = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "MARCA VEHICULO"
                vPosicion = 159
                vgMARCADEVEHICULO = Mid(sLine, vPosicion, 30)
                '      -------------------------------------------------------
                vCampo = "MODELO"
                vPosicion = 189
                vgMODELO = Mid(sLine, vPosicion, 20)
                '      -------------------------------------------------------
                'vCampo = "COLOR"
                'vPosicion = 209
                'vgMARCADEVEHICULO = Mid(sLine, 194, 30)
                '      -------------------------------------------------------
                vCampo = "AÑO"
                vPosicion = 239
                vgAno = Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "PATENTE"
                vPosicion = 243
                vgPATENTE = Mid(sLine, vPosicion, 15)
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 258
                'vgPATENTE = Mid(sLine, 263, 8)
                '      -------------------------------------------------------
                vCampo = "COB. AUTO"
                vPosicion = 260
                'vgCOBERTURAVEHICULO = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                'vCampo = "COB. VIAJ."
                'vPosicion = 261
                'vgCOBERTURAVIAJERO = Mid(sLine, 279, 1)
                '      -------------------------------------------------------
                'vCampo = "COB. HOGAR"
                'vPosicion = 262
                'vgCOBERTURAHOGAR = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                vCampo = "PLAN COMERCIAL"
                vPosicion = 263
                vgCodigoDeServicioVip = Trim(Mid(sLine, vPosicion, 5))
                If vgCodigoDeServicioVip = "99_GE" Or vgCodigoDeServicioVip = "50_GZ" Then
                    vgCodigoDeServicioVip = "1"
                Else
                    vgCodigoDeServicioVip = "0"
                End If
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 268
                '      -------------------------------------------------------
                vCampo = "PREST.ADIC."
                vPosicion = 274
                vgCargo = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                'vCampo = "COD.PROC."
                'vPosicion = 275
                'vgCodigoDeProductor = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                vCampo = "CRISTALES"
                vPosicion = 275
                vgInformacionAdicionalValor1 = Mid(sLine, vPosicion, 15) ' 12 enteros + coma + 2 decimales
                vgInformacionAdicionalValor1 = Replace(vgInformacionAdicionalValor1, ".", ",") ' en subdesarrollo, el double es con "," y no con "."
                vgInformacionAdicionalValor1 = CStr(CDbl(vgInformacionAdicionalValor1))
                '      -------------------------------------------------------
                vCampo = "LUNETA/PARABRISAS"
                vPosicion = 290
                vgInformacionAdicionalValor2 = Mid(sLine, vPosicion, 15) ' 12 enteros + coma + 2 decimales
                vgInformacionAdicionalValor2 = Replace(vgInformacionAdicionalValor2, ".", ",") ' en subdesarrollo, el double es con "," y no con "."
                vgInformacionAdicionalValor2 = CStr(CDbl(vgInformacionAdicionalValor2))
                '      -------------------------------------------------------
                vCampo = "TECHO"
                vPosicion = 305
                vgInformacionAdicionalValor3 = Mid(sLine, vPosicion, 15) ' 12 enteros + coma + 2 decimales
                vgInformacionAdicionalValor3 = Replace(vgInformacionAdicionalValor3, ".", ",") ' en subdesarrollo, el double es con "," y no con "."
                vgInformacionAdicionalValor3 = CStr(CDbl(vgInformacionAdicionalValor3))
                '      -------------------------------------------------------
                vCampo = "CERRADURA"
                vPosicion = 320
                vgInformacionAdicionalValor4 = Mid(sLine, vPosicion, 15) ' 12 enteros + coma + 2 decimales
                vgInformacionAdicionalValor4 = Replace(vgInformacionAdicionalValor4, ".", ",") ' en subdesarrollo, el double es con "," y no con "."
                vgInformacionAdicionalValor4 = CStr(CDbl(vgInformacionAdicionalValor4))
                '      -------------------------------------------------------
                vCampo = "COBERTURA"
                vPosicion = 335
                'vgCOBERTURAVEHICULO = Mid(sLine, vPosicion, 3)
    
            '   Vida colectivo
            ElseIf vgidCampana = 1075 Then
                vCampo = "NRO. EMPRESA"
                vPosicion = 1
                vgNUMEROCOMPANIA = Mid(sLine, vPosicion, 3)
                '      -------------------------------------------------------
                vCampo = "NRO. POLIZA"
                vPosicion = 4
                vgNROPOLIZA = Val(Trim(Mid(sLine, vPosicion, 20)))
                '      -------------------------------------------------------
                vCampo = "ITEM"
                vPosicion = 24
                vgNROSECUENCIAL = Mid(sLine, vPosicion, 6)
                vgNROPOLIZA = Trim(vgNROPOLIZA) & Trim(vgNROSECUENCIAL)
                '      -------------------------------------------------------
                vCampo = "APELLIDOYNOMBRE"
                vPosicion = 30
                vgAPELLIDOYNOMBRE = Mid(sLine, vPosicion, 50)
                '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 80
                vgDOMICILIO = Mid(sLine, vPosicion, 55)
                '      -------------------------------------------------------
                vCampo = "COD. POSTAL"
                vPosicion = 135
                vgCODIGOPOSTAL = Mid(sLine, vPosicion, 8)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA DESDE"
                vPosicion = 143
                vgFECHAVIGENCIA = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "VIGENCIA HASTA"
                vPosicion = 151
                vgFECHAVENCIMIENTO = Mid(sLine, vPosicion + 6, 2) & "/" & Mid(sLine, vPosicion + 4, 2) & "/" & Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                'vCampo = "MARCA VEHICULO"
                'vPosicion = 159
                'vgMARCADEVEHICULO = Mid(sLine, vPosicion, 30)
                '      -------------------------------------------------------
                'vCampo = "MODELO"
                'vPosicion = 189
                'vgMODELO = Mid(sLine, vPosicion, 20)
                '      -------------------------------------------------------
                'vCampo = "COLOR"
                'vPosicion = 209
                'vgMARCADEVEHICULO = Mid(sLine, 194, 30)
                '      -------------------------------------------------------
                vCampo = "AÑO"
                vPosicion = 239
                vgTipodeDocumento = Mid(sLine, vPosicion, 4)
                '      -------------------------------------------------------
                vCampo = "PATENTE"
                vPosicion = 243
                vgNumeroDeDocumento = Val(Trim(Mid(sLine, vPosicion, 15)))
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 258
                'vgPATENTE = Mid(sLine, 263, 8)
                '      -------------------------------------------------------
                'vCampo = "COB. AUTO"
                'vPosicion = 260
                'vgCOBERTURAVEHICULO = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                vCampo = "COB. VIAJ."
                vPosicion = 261
                'vgCOBERTURAVIAJERO = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                'vCampo = "COB. HOGAR"
                'vPosicion = 262
                'vgCOBERTURAHOGAR = Mid(sLine, vPosicion, 1)
                '      -------------------------------------------------------
                'vCampo = "SIN USO"
                'vPosicion = 263
                '      -------------------------------------------------------
                'vCampo = "PREST.ADIC."
                'vPosicion = 273
                'vgPrestador = Mid(sLine, vPosicion, 1)   'no existe esta variable
                '      -------------------------------------------------------
                'vCampo = "COD.PROC."
                'vPosicion = 332
                'vgCodigoDeProductor = Mid(sLine, 332, 5)
    
            End If
            
            sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & vgidCampana & "  and idproductoencliente = '" & vgCodigoEnCliente & "'"
            rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly
            If Not rsprod.EOF Then
                vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                
                If vgCOBERTURAVEHICULO <> "" Then
                    vLeidosPorCoberturaLista(0) = vLeidosPorCoberturaLista(0) + 1
                End If
                
                vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                
                If vgCOBERTURAVIAJERO <> "" Then
                    vLeidosPorCoberturaLista(1) = vLeidosPorCoberturaLista(1) + 1
                End If
                
                vgCOBERTURAHOGAR = rsprod("coberturahogar")
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                
                If vgCOBERTURAHOGAR <> "" Then
                    vLeidosPorCoberturaLista(2) = vLeidosPorCoberturaLista(2) + 1
                End If
                
            Else
                vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", lRow, sName)
            End If
            rsprod.Close
            
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
            Dim vcamp As Integer
            Dim vdif As Long
            
            ssql = "select top(1) "
            ssql = ssql & " APELLIDOYNOMBRE,"
            ssql = ssql & " DOMICILIO,"
            ssql = ssql & " CODIGOPOSTAL,"
            ssql = ssql & " FECHAVIGENCIA,"
            ssql = ssql & " FECHAVENCIMIENTO,"
            ssql = ssql & " FECHABAJAOMNIA,"
            ssql = ssql & " MARCADEVEHICULO,"
            ssql = ssql & " MODELO,"
            ssql = ssql & " Color,"
            ssql = ssql & " ANO,"
            ssql = ssql & " PATENTE,"
            ssql = ssql & " TipoDeDocumento,"
            ssql = ssql & " NumeroDeDocumento,"
            ssql = ssql & " COBERTURAVEHICULO,"
            ssql = ssql & " COBERTURAVIAJERO,"
            ssql = ssql & " COBERTURAHOGAR,"
            ssql = ssql & " Cargo,"
            ssql = ssql & " MontoCoverturaVidrios,"

            If vgidCampana <> 1077 Then
                ssql = ssql & " idpoliza"
                ssql = ssql & " from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            Else
                ssql = ssql & " p.idpoliza,"
                ssql = ssql & " CodigoDeServicioVip,"
                ssql = ssql & " Valor1,"
                ssql = ssql & " Valor2,"
                ssql = ssql & " Valor3,"
                ssql = ssql & " Valor4"
                ssql = ssql & " from Auxiliout.dbo.tm_Polizas p"
                ssql = ssql & " left join TM_InformacionAdicionalPorPoliza iap on iap.IDPOLIZA = p.IDPOLIZA"
                ssql = ssql & " where IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            End If

            
            'ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly ' rscn1 no hace conexion con ADOB.conection
            
            If Err Then
                MsgBox Err.Description
            End If
            
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1 ' REVISAR
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1 ' no se recibe en ninguna de las bases
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("TipoDeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Val(Trim(rscn1("COBERTURAVEHICULO"))) <> Val(Trim(vgCOBERTURAVEHICULO)) Then vdif = vdif + 1
                If Val(Trim(rscn1("COBERTURAVIAJERO"))) <> Val(Trim(vgCOBERTURAVIAJERO)) Then vdif = vdif + 1
                If Val(Trim(rscn1("COBERTURAHOGAR"))) <> Val(Trim(vgCOBERTURAHOGAR)) Then vdif = vdif + 1
                If Trim(rscn1("Cargo")) <> Trim(vgCargo) Then vdif = vdif + 1
                
                If vgidCampana = 1076 Then ' solo para HOGAR/INTEGRAL
                    If (Trim(rscn1("MontoCoverturaVidrios")) <> Trim(vgTopeCritales)) Or IsNull(rscn1("MontoCoverturaVidrios")) Then vdif = vdif + 1
                End If
                
                If vgidCampana <> 1077 Then
                    vgIDPOLIZA = rscn1("idpoliza")
                Else
                    If IsNull(rscn1("CodigoDeServicioVip")) Then vdif = vdif + 1
                    If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                    If (Trim(rscn1("Valor1")) <> Trim(vgInformacionAdicionalValor1)) Or IsNull(rscn1("Valor1")) Then vdif = vdif + 1
                    If Trim(rscn1("Valor2")) <> Trim(vgInformacionAdicionalValor2) Then vdif = vdif + 1
                    If Trim(rscn1("Valor3")) <> Trim(vgInformacionAdicionalValor3) Then vdif = vdif + 1
                    If Trim(rscn1("Valor4")) <> Trim(vgInformacionAdicionalValor4) Then vdif = vdif + 1

                    vgIDPOLIZA = rscn1("idpoliza")
                End If

            End If
            
            If vdif > 0 Then
                vdif = vdif
            End If
            
            
            rscn1.Close
            '-=================================================================================================================
            
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos1073("
            ssql = ssql & "IDPOLIZA, "
            ssql = ssql & "IDCIA, "
            ssql = ssql & "NUMEROCOMPANIA, "
            ssql = ssql & "NROPOLIZA, "
            'ssql = ssql & "NROSECUENCIAL, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "CODIGOPOSTAL, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "ANO, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "Cargo, "
            ssql = ssql & "MontoCoverturaVidrios, "
            ssql = ssql & "CodigoDeServicioVip, "
            ssql = ssql & "InformacionAdicionalValor1, "
            ssql = ssql & "InformacionAdicionalValor2, "
            ssql = ssql & "InformacionAdicionalValor3, "
            ssql = ssql & "InformacionAdicionalValor4, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones )"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            'ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgAno) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgCodigoEnCliente) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", "
            ssql = ssql & Trim(vgidCampana) & ", '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgCargo) & "', '"        'Prestador
            ssql = ssql & Trim(vgTopeCritales) & "', '"
            ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
            ssql = ssql & Trim(vgInformacionAdicionalValor1) & "', '"
            ssql = ssql & Trim(vgInformacionAdicionalValor2) & "', '"
            ssql = ssql & Trim(vgInformacionAdicionalValor3) & "', '"
            ssql = ssql & Trim(vgInformacionAdicionalValor4) & "', '"
            ssql = ssql & Trim(vLote) & "', '"
            ssql = ssql & Trim(vdif) & "') "
            cn.Execute ssql
            
            '========Control de errores=========================================================
            If Err Then
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, "")
                Err.Clear
            End If
            '===================================================================================
            
            If vdif > 0 Then
                regMod = regMod + 1
            End If
            
            Ll = Ll + 1
            ll100 = ll100 + 1
            If ll100 = 100 Then
                'ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
                ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " Archivo " & fileimportacion & Chr(13) & " copiando linea " & Ll & Chr(13) & " Total de registros: " & totalRegistros & Chr(13)
                
                ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & vgidCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
                
                ll100 = 0
            End If
            DoEvents
    
        Loop
        
        If vLeidosPorCoberturaLista(0) > 0 Then
            CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAVEHICULO", vgCOBERTURAVEHICULO, vLeidosPorCoberturaLista(0), 0
        End If
        If vLeidosPorCoberturaLista(1) > 0 Then
            CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAVIAJERO", vgCOBERTURAVIAJERO, vLeidosPorCoberturaLista(1), 0
        End If
        If vLeidosPorCoberturaLista(2) > 0 Then
            CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAHOGAR", vgCOBERTURAHOGAR, vLeidosPorCoberturaLista(2), 0
        End If
        vLeidosPorCoberturaLista(0) = 0
        vLeidosPorCoberturaLista(1) = 0
        vLeidosPorCoberturaLista(2) = 0
        
        If finProcesoCampana = 1 Then
            Select Case vgidCampana
                Case vidCampanaLista(0)
                    vLotesLista(0) = vLote
                    vLeidosLista(0) = Ll
                Case vidCampanaLista(1)
                    vLotesLista(1) = vLote
                    vLeidosLista(1) = Ll
                Case vidCampanaLista(2)
                    vLotesLista(2) = vLote
                    vLeidosLista(2) = Ll
                Case vidCampanaLista(3)
                    vLotesLista(3) = vLote
                    vLeidosLista(3) = Ll
            End Select
            finProcesoCampana = 0
        End If
        
    End If

Next

If archivoLeido = 0 Then
    MsgBox "No se encontro el archivo"
    Exit Sub
End If


'================Control de Leidos===============================================
Dim cont As Integer
For cont = 0 To 3
    If vLotesLista(cont) <> 0 Then
        cn1.Execute "TM_CargaPolizasLogDeSetLeidosXCampana " & vgCORRIDA & ", " & vidCampanaLista(cont) & ", " & vLeidosLista(cont)
    End If
Next cont

'cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & vLeidosLista(0) 'SOLO PARA ATM
listoParaProcesar
'=================================================================================

ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub


'===============Inicio del Control de Procesos=====================================
cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
Dim rsCMP As New Recordset
DoEvents

'Dim cont As Integer
For cont = 0 To 3
    If vLotesLista(cont) <> 0 Then
        For lLote = 1 To vLotesLista(cont)
            cn1.CommandTimeout = 300
            cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & vidCampanaLista(cont)
            ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vidCampanaLista(cont)
            rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            ImportadordePolizas.txtprocesando.Text = "idCampana " & vidCampanaLista(cont) & Chr(13) & "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vLeidosLista(cont) & " Procesando los datos"
            DoEvents
            rsCMP.Close
        Next lLote
    End If
Next cont

'================Baja de polizas===============================================
For cont = 0 To 3
    If vLotesLista(cont) <> 0 Then
        cn1.Execute "TM_BajaDePolizasSegurcoopControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vidCampanaLista(cont)
    End If
Next cont

'============Finaliza Proceso========================================================
For cont = 0 To 3
    If vLotesLista(cont) <> 0 Then
        cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & vidCampanaLista(cont) & ", " & vgCORRIDA
        CoberturasProcesadas vgCORRIDA, vidCampanaLista(cont), vidHistorialImportacionLista(cont) ' NUEVO
        CoberturasBajas vgCORRIDA, vidCampanaLista(cont), vidHistorialImportacionLista(cont) ' NUEVO
        Procesado
    End If
Next cont

'cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
'Procesado
'=====================================================================================
ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " FinDeProceso"
ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF

vgidCampana = 1073

Exit Sub
errores:
vgErrores = 1
If Ll = 0 Then
    MsgBox Err.Description
Else
    MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
End If

End Sub

