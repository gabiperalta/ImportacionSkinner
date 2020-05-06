Attribute VB_Name = "triunfoSeguros"
Option Explicit
Public Sub ImportarTriunfoSeguros()
Dim gsServidor As String, gsBaseEmpresa As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
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

Dim sAno As Date
Dim sMes As Date
Dim sDia As Date
Dim sZone As String
Dim sArchivo As String
Dim sExtencion As String

'Dim vLeidosLista(0 To 3) As Long
Dim vCoberturaLista(0 To 2, 0 To 20) As String            ' campana,tipoCobertura,coberturaEncontrada
Dim vLeidosPorCoberturaLista(0 To 2, 0 To 20) As Long     ' campana,tipoCobertura,coberturaEncontrada
'Dim vLeidosPorCoberturaLista(0 To 2) As Long                       ' tipoCobertura
Dim vCoberturaActual(0 To 2) As String

cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos2"
On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

Dim vCantDeErrores As Long
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores" 'creacion de archivo . log que almacena los errores surgidos de la importacion ( registros no importados)
vCantDeErrores = 0

If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

Ll = 0
ll100 = 0
regMod = 0

nroArchivo = 0
archivoLeido = 0
sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & "TriunfoSeguros.txt"
'For nroArchivo = 1 To 2

For i = 0 To 20
    vCoberturaLista(0, i) = "_"
    vCoberturaLista(1, i) = "_"
    vCoberturaLista(2, i) = "_"
Next i


Do While fs.FileExists(sFile)
    
    archivoLeido = 1

    Set tf = fs.OpenTextFile(sFile, ForReading, True)

'=======control de lectura del archivo de datos=======================
    If Err Then
        MsgBox Err.Description & " en linea: " & Ll
        Err.Clear
        Exit Sub
    End If
    
    
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
    
    End If
    
    Do Until tf.AtEndOfStream
    
        Blanquear
        vCoberturaActual(0) = ""
        vCoberturaActual(1) = ""
        vCoberturaActual(2) = ""
    
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
        'If Trim(Mid(sLine, 271, 4)) = "A" Or Trim(Mid(sLine, 271, 4)) = "V" Then
        '    vTipodeServicioActual = "AV"
        'Else
        '    vTipodeServicioActual = "H"
        'End If
        'If vLote = 1 And nroLinea = 1 Then
        '    vTipodeServicio = vTipodeServicioActual
        'Else
        '    If vTipodeServicio <> vTipodeServicioActual Then
        '        MsgBox "Error En Base Imposible Importar, Tipo de Servicio Incorrecto en Linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion, , "Control de tipo de Base"
        '        Exit Sub
        '    End If
        'End If
        
'=======Control de Lote===============================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
              vLote = vLote + 1
            nroLinea = 1
        End If
'=====================================================
        
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

    If Trim(Mid(sLine, 271, 4)) = "A" Or Trim(Mid(sLine, 271, 4)) = "V" Then
    '        vgIDPOLIZA = Mid(sLine, 1, 10)
    '        vgIDCIA = Mid(sLine, 1, 10)
     '      -------------------------------------------------------
            vCampo = "Compania"
            vPosicion = 1
            vgNUMEROCOMPANIA = Mid(sLine, 1, 3)
     '      -------------------------------------------------------
            vCampo = "NroPoliza"
            vPosicion = 4
            vgNROPOLIZA = Mid(sLine, 4, 13)
     '      -------------------------------------------------------
            vCampo = "NroSecuencial"
            vPosicion = 17
            vgNROSECUENCIAL = Mid(sLine, 17, 3)
     '      -------------------------------------------------------
            vCampo = "APELLIDOYNOMBRE"
            vPosicion = 20
            vgAPELLIDOYNOMBRE = Mid(sLine, 20, 50)
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
            If Trim(vgTipodeServicio) <> "A" And Trim(vgTipodeServicio) <> "V" Then
                MsgBox "Error En Base Imposible Importar, Tipo de Servicio Incorrecto en Linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion, , "Control de tipo de Base"
                Exit Sub
            End If
     '      -------------------------------------------------------
            vCampo = "TIPODEVEHICULO"
            vPosicion = 275
            Select Case Trim(Mid(sLine, 275, 2))
                Case "A"
                    vgTIPODEVEHICULO = 1
                Case "T"
                    vgTIPODEVEHICULO = 2
                Case "C"
                    vgTIPODEVEHICULO = 4
                Case "M"
                    vgTIPODEVEHICULO = 5
                Case "P"
                    vgTIPODEVEHICULO = 3
                Case Else
                    vgTIPODEVEHICULO = 0
            End Select
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
            vCampo = "InformadoSinCobertura"
            vPosicion = 338
            vgInformadoSinCobertura = Mid(sLine, 338, 1)
            If vgInformadoSinCobertura = "N" Then
                vgCOBERTURAVEHICULO = 98
                vgCOBERTURAVIAJERO = 98
            End If
     '      -------------------------------------------------------
    '        vgTipodeDocumento = Mid(sLine, 1, 10)
    '        vgNumeroDeDocumento = Mid(sLine, 1, 10)
    '        vgTipodeHogar = Mid(sLine, 1, 10)
    '        vgIniciodeAnualidad = Mid(sLine, 1, 10)
    '        vgPolizaIniciaAnualidad = Mid(sLine, 1, 10)
    '        vgTelefono = Mid(sLine, 1, 10)
    '        vgNroMotor = Mid(sLine, 1, 10)
    '        vgGama = Mid(sLine, 1, 10)
    ElseIf Trim(Mid(sLine, 271, 4)) = "H" Then

    '        vgIDPOLIZA = Mid(sLine, 1, 10)
    '        vgIDCIA = Mid(sLine, 1, 10)
     '      -------------------------------------------------------
            vCampo = "Compania"
            vPosicion = 1
            vgNUMEROCOMPANIA = Mid(sLine, 1, 3)
     '      -------------------------------------------------------
            vCampo = "NroPoliza"
            vPosicion = 4
            vgNROPOLIZA = Mid(sLine, 4, 13)
     '      -------------------------------------------------------
            vCampo = "NroSecuencial"
            vPosicion = 17
            vgNROSECUENCIAL = Mid(sLine, 17, 3)
     '      -------------------------------------------------------
            vCampo = "APELLIDOYNOMBRE"
            vPosicion = 20
            vgAPELLIDOYNOMBRE = Mid(sLine, 20, 50)
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
'            vCampo = "MARCADEVEHICULO"
'            vPosicion = 194
'            vgMARCADEVEHICULO = Mid(sLine, 194, 30)
     '      -------------------------------------------------------
'            vCampo = "MODELO"
'            vPosicion = 224
'            vgMODELO = Mid(sLine, 224, 20)
     '      -------------------------------------------------------
            vCampo = "Numero De Documento"
            vPosicion = 244
            vgNumeroDeDocumento = Mid(sLine, 244, 15)
     '      -------------------------------------------------------
'            vCampo = "AÑO"
'            vPosicion = 259
'            vgANO = Mid(sLine, 259, 4)
'     '      -------------------------------------------------------
'            vCampo = "PATENTE"
'            vPosicion = 263
'            vgPATENTE = Mid(sLine, 263, 8)
     '      -------------------------------------------------------
            vCampo = "TipodeServicio"
            vPosicion = 271
            vgTipodeServicio = Mid(sLine, 271, 4)
            If Trim(vgTipodeServicio) <> "H" Then
                MsgBox "Error En Base Imposible Importar, Tipo de Servicio Incorrecto en Linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion, , "Control de tipo de Base"
                Exit Sub
            End If
     '      -------------------------------------------------------
            vCampo = "TIPO DE Hogar"
            vPosicion = 275
            vgTipodeHogar = Trim(Mid(sLine, 275, 2))
            Select Case vgTipodeHogar
            
            Case "1"
                vgCOBERTURAHOGAR = "01"
            Case "2"
                vgCOBERTURAHOGAR = "02"
            Case "3"
                vgCOBERTURAHOGAR = "03"
            End Select
            
     '      -------------------------------------------------------
    '        vgIDTIPODECOBERTURA = Mid(sLine, 1, 10)
     '      -------------------------------------------------------
'            vCampo = "COBERTURAVEHICULO"
'            vPosicion = 277
'            vgCOBERTURAVEHICULO = Mid(sLine, 277, 2)
'     '      ------------------------------------------------------
'            vCampo = "COBERTURAVIAJERO"
'            vPosicion = 279
'            vgCOBERTURAVIAJERO = Mid(sLine, 279, 2)
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
'            vCampo = "Conductor"
'            vPosicion = 282
'            vgConductor = Mid(sLine, 282, 50)
     '      -------------------------------------------------------
            vCampo = "CodigoDeProductor"
            vPosicion = 332
            vgCodigoDeProductor = Mid(sLine, 332, 5)
     '      -------------------------------------------------------
            vCampo = "CodigoDeServicioVip"
            vPosicion = 337
            vgCodigoDeServicioVip = Mid(sLine, 337, 1)
     '      -------------------------------------------------------
            vCampo = "InformadoSinCobertura"
            vPosicion = 338
            vgInformadoSinCobertura = Mid(sLine, 338, 1)
            If vgInformadoSinCobertura = "N" Then
                vgCOBERTURAVEHICULO = 0
                vgCOBERTURAHOGAR = 98
                vgCOBERTURAVIAJERO = 0
            End If
     '      -------------------------------------------------------
            vCampo = "Tope Cristales"
            vPosicion = 339
            vgTopeCritales = Mid(sLine, 339, 7)
     '      -------------------------------------------------------
    '        vgTipodeDocumento = Mid(sLine, 1, 10)
    '        vgNumeroDeDocumento = Mid(sLine, 1, 10)
    '        vgTipodeHogar = Mid(sLine, 1, 10)
    '        vgIniciodeAnualidad = Mid(sLine, 1, 10)
    '        vgPolizaIniciaAnualidad = Mid(sLine, 1, 10)
    '        vgTelefono = Mid(sLine, 1, 10)
    '        vgNroMotor = Mid(sLine, 1, 10)
    '        vgGama = Mid(sLine, 1, 10)
    
    
    End If
    
    '=============== Lectura de coberturas ===============
    vCoberturaActual(0) = vgCOBERTURAVEHICULO
    vCoberturaActual(1) = vgCOBERTURAVIAJERO
    vCoberturaActual(2) = vgCOBERTURAHOGAR
    
    
    Dim posicion As Integer
    'Dim encontrado As Integer
    Dim coberturaPosicion As Integer
    
    'encontrado = 0
        
    For coberturaPosicion = 0 To 2
        posicion = 0
        Do While posicion < 20 And vCoberturaActual(coberturaPosicion) <> ""
            If vCoberturaLista(coberturaPosicion, posicion) = vCoberturaActual(coberturaPosicion) Then
                vLeidosPorCoberturaLista(coberturaPosicion, posicion) = vLeidosPorCoberturaLista(coberturaPosicion, posicion) + 1
                Exit Do
            ElseIf vCoberturaLista(coberturaPosicion, posicion) = "_" Then
                vCoberturaLista(coberturaPosicion, posicion) = vCoberturaActual(coberturaPosicion)
                vLeidosPorCoberturaLista(coberturaPosicion, posicion) = vLeidosPorCoberturaLista(coberturaPosicion, posicion) + 1
                Exit Do
            End If
            
            posicion = posicion + 1
        Loop
    
'        For posicion = 0 To 20
'            If vLeidosPorCoberturaLista(coberturaPosicion, posicion) = vCoberturaActual(coberturaPosicion) Then
'                encontrado = 1
'                Exit For
'            End If
'        Next posicion
'
'        If encontrado = 0 Then
'            For posicion = 0 To 20
'                If vidCampanaLista(posicion) = 0 Then
'                    vidCampanaLista(posicion) = vgidCampana
'                    Exit For
'                End If
'            Next posicion
'        End If
    Next coberturaPosicion
    
    
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
    ssql = ssql & " PROVINCIA,"
    ssql = ssql & " CODIGOPOSTAL, "
    ssql = ssql & " FECHAVENCIMIENTO,"
    ssql = ssql & " FECHABAJAOMNIA,"
    ssql = ssql & " IDAUTO,"
    ssql = ssql & " MARCADEVEHICULO,"
    ssql = ssql & " MODELO,"
    ssql = ssql & " COLOR,"
    ssql = ssql & " PATENTE,"
    ssql = ssql & " TIPODEVEHICULO,"
    ssql = ssql & " TipodeServicio,"
    ssql = ssql & " COBERTURAVEHICULO,"
    ssql = ssql & " COBERTURAVIAJERO,"
    ssql = ssql & " LOCALIDAD,"
    ssql = ssql & " TipodeOperacion,"
    ssql = ssql & " Operacion,"
    ssql = ssql & " CATEGORIA,"
    ssql = ssql & " ASISTENCIAXENFERMEDAD,"
    ssql = ssql & " Conductor,"
    ssql = ssql & " CodigoDeProductor,"
    ssql = ssql & " CodigoDeServicioVip,"
    ssql = ssql & " TipodeDocumento,"
    ssql = ssql & " NumeroDeDocumento,"
    ssql = ssql & " TipodeHogar,"
    ssql = ssql & " IniciodeAnualidad,"
    ssql = ssql & " PolizaIniciaAnualidad,"
    ssql = ssql & " idpoliza"
    ssql = ssql & " from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' and Nrosecuencial = '" & vgNROSECUENCIAL & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly ' rscn1 no hace conexion con ADOB.conection
    
    If Err Then
        MsgBox Err.Description & " en linea: " & Ll
    End If
    
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
    vgCambiosDetectados = ""
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("APELLIDOYNOMBRE")) & "<>" & Trim(vgAPELLIDOYNOMBRE) & ", "
                End If
                
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("DOMICILIO")) & "<>" & Trim(vgDOMICILIO) & ", "
                End If
                
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("LOCALIDAD")) & "<>" & Trim(vgLOCALIDAD) & ", "
                End If
                
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("PROVINCIA")) & "<>" & Trim(vgPROVINCIA) & ", "
                End If
                
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("CODIGOPOSTAL")) & "<>" & Trim(vgCODIGOPOSTAL) & ", "
                End If
                
'                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1

                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("FECHAVENCIMIENTO")) & "<>" & Trim(vgFECHAVENCIMIENTO) & ", "
                End If
                
                If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("FECHABAJAOMNIA")) & "<>" & Trim(vgFECHABAJAOMNIA) & ", "
                End If
                
                If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("IDAUTO")) & "<>" & Trim(vgIDAUTO) & ", "
                End If
                
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("MARCADEVEHICULO")) & "<>" & Trim(vgMARCADEVEHICULO) & ", "
                End If
                
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("MODELO")) & "<>" & Trim(vgMODELO) & ", "
                End If
                
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("COLOR")) & "<>" & Trim(vgCOLOR) & ", "
                End If
                
'                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1

                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("PATENTE")) & "<>" & Trim(vgPATENTE) & ", "
                End If
                
                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("TIPODEVEHICULO")) & "<>" & Trim(vgTIPODEVEHICULO) & ", "
                End If
                
                If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("Tip odeServicio")) & "<>" & Trim(vgTipodeServicio) & ", "
                End If
                
'                If Trim(rscn1("IDTIPODECOBERTURA")) <> Trim(vgIDTIPODECOBERTURA) Then vdif = vdif + 1

                If Val(Trim(rscn1("COBERTURAVEHICULO"))) <> Val(Trim(vgCOBERTURAVEHICULO)) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("COBERTURAVEHICULO")) & "<>" & Trim(vgCOBERTURAVEHICULO) & ", "
                End If
                
                If Val(Trim(rscn1("COBERTURAVIAJERO"))) <> Val(Trim(vgCOBERTURAVIAJERO)) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("LOCALIDAD")) & "<>" & Trim(vgLOCALIDAD) & ", "
                End If
                
                If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("TipodeOperacion")) & "<>" & Trim(vgTipodeOperacion) & ", "
                End If
                
                If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("Operacion")) & "<>" & Trim(vgOperacion) & ", "
                End If
                
                If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("CATEGORIA")) & "<>" & Trim(vgCATEGORIA) & ", "
                End If
                
                If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("ASISTENCIAXENFERMEDAD")) & "<>" & Trim(vgASISTENCIAXENFERMEDAD) & ", "
                End If
                
'                If Trim(rscn1("IdCampana")) <> Trim(vgIdCampana) Then vdif = vdif + 1

                If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("Conductor")) & "<>" & Trim(vgConductor) & ", "
                End If
                
                
                If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("CodigoDeProductor")) & "<>" & Trim(vgCodigoDeProductor) & ", "
                End If
                
                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("CodigoDeServicioVip")) & "<>" & Trim(vgCodigoDeServicioVip) & ", "
                End If
                
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("TipodeDocumento")) & "<>" & Trim(vgTipodeDocumento) & ", "
                End If
                
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("NumeroDeDocumento")) & "<>" & Trim(vgNumeroDeDocumento) & ", "
                End If
                
                If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("TipodeHogar")) & "<>" & Trim(vgTipodeHogar) & ", "
                End If
                
                If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("IniciodeAnualidad")) & "<>" & Trim(vgIniciodeAnualidad) & ", "
                End If
                
                If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then
                    vdif = vdif + 1
                    vgCambiosDetectados = vgCambiosDetectados & " " & Trim(rscn1("PolizaIniciaAnualidad")) & "<>" & Trim(vgPolizaIniciaAnualidad) & ", "
                End If
                
                vgIDPOLIZA = rscn1("idpoliza")
            End If
            
            If vdif > 0 Then
                vdif = vdif
            End If
            

        rscn1.Close
'-=================================================================================================================
  
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos2 ("
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
        ssql = ssql & "IdLote, "
        ssql = ssql & "InformadoSinCobertura, "
        ssql = ssql & "MontoCoverturaVidrios, "
        ssql = ssql & "Modificaciones, "
        ssql = ssql & " CambiosDetectados )"
        
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
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
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
        ssql = ssql & Trim(vgInformadoSinCobertura) & "', '"
        ssql = ssql & Trim(vgTopeCritales) & "', '"
        ssql = ssql & Trim(vdif) & "',' "
        ssql = ssql & Trim(vgCambiosDetectados) & "') "
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

    
    nroArchivo = nroArchivo + 1
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & "TriunfoSeguros" & nroArchivo & ".txt"
    
'Next nroArchivo
Loop

If archivoLeido = 0 Then
    MsgBox "No se encontro el archivo"
    Exit Sub
End If

For posicion = 0 To 20

    If vLeidosPorCoberturaLista(0, posicion) > 0 Then
        CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAVEHICULO", vCoberturaLista(0, posicion), vLeidosPorCoberturaLista(0, posicion), 0
    End If
    If vLeidosPorCoberturaLista(1, posicion) > 0 Then
        CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAVIAJERO", vCoberturaLista(1, posicion), vLeidosPorCoberturaLista(1, posicion), 0
    End If
    If vLeidosPorCoberturaLista(2, posicion) > 0 Then
        CantidadPorCobertura vgIdHistorialImportacion, "COBERTURAHOGAR", vCoberturaLista(2, posicion), vLeidosPorCoberturaLista(2, posicion), 0
    End If

Next posicion

      
'================Control de Leidos===============================================
                            cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
                            listoParaProcesar
'=================================================================================
        
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

'Exit Sub

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
    
'================Baja de polizas===============================================
    cn1.Execute "TM_BajaDePolizasHogarYAutos" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

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
