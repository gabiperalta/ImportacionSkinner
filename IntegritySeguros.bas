Attribute VB_Name = "IntegritySeguros"
Option Explicit
Public Sub ImportarIntegritySeguros()


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
Dim sArchivo As String
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long

'On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos347"
    vIDCIA = 9999728
    vIDCampana = 347
    LongDeLote = 1000

    Ll = 0
    'Manejar el tema aqui para seleccionar que base viene
    sArchivo = Mid(FileImportacion, 1, InStr(1, FileImportacion, "/") - 1)
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo
    vidTipoDePoliza = 1
    If Not fs.FileExists(sFile) Then
        sArchivo = Mid(FileImportacion, InStr(1, FileImportacion, "/") + 1)
        sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo
        vidTipoDePoliza = 3
        If Not fs.FileExists(sFile) Then
            vidTipoDePoliza = 0
            Exit Sub
        End If
    End If
    FileImportacion = sArchivo
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
        If nroLinea = 465 Then
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
        If vidTipoDePoliza = 1 Then
            vTipoDePoliza = "Vehiculo"
         '      -------------------------------------------------------
                vCampo = "PATENTE"
                vPosicion = 1
                vgPATENTE = Mid(sLine, 1, 9)
         '      -------------------------------------------------------
                vCampo = "MOTOR"
                vPosicion = 10
                vgNroMotor = Mid(sLine, 10, 20)
         '      -------------------------------------------------------
                vCampo = "POLIZA"
                vPosicion = 30
                vgNROPOLIZA = Mid(sLine, 30, 15)
         '      -------------------------------------------------------
                vCampo = "APELLIDO Y NOMBRE CLIENTE"
                vPosicion = 45
                vgAPELLIDOYNOMBRE = Mid(sLine, 45, 50)
         '      -------------------------------------------------------
                vCampo = "TIPO DOCUMENTO"
                vPosicion = 95
                If Trim(Mid(sLine, 95, 2)) = "00" Then
                    vgTipodeDocumento = "CI Policia Federal"
                ElseIf Trim(Mid(sLine, 95, 2)) = "80" Then
                    vgTipodeDocumento = "CUIT"
                ElseIf Trim(Mid(sLine, 95, 2)) = "89" Then
                    vgTipodeDocumento = "LE"
                ElseIf Trim(Mid(sLine, 95, 2)) = "90" Then
                    vgTipodeDocumento = "LC"
                ElseIf Trim(Mid(sLine, 95, 2)) = "94" Then
                    vgTipodeDocumento = "Pasaporte"
                ElseIf Trim(Mid(sLine, 95, 2)) = "96" Then
                    vgTipodeDocumento = "DNI"
                ElseIf Trim(Mid(sLine, 95, 2)) = "99" Then
                    vgTipodeDocumento = "CUIT Externo"
                ElseIf Trim(Mid(sLine, 95, 2)) = "82" Then
                    vgTipodeDocumento = "CUIL"
                End If
         '      -------------------------------------------------------
                vCampo = "NUMERO DOCUMENTO"
                vPosicion = 97
                vgNumeroDeDocumento = Mid(sLine, 97, 11)
         '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 108
                vgDOMICILIO = Mid(sLine, 108, 35)
         '      -------------------------------------------------------
                vCampo = "CODIGO POSTAL"
                vPosicion = 143
                vgCODIGOPOSTAL = Mid(sLine, 143, 5)
         '      -------------------------------------------------------
                vCampo = "LOCALIDAD"
                vPosicion = 148
                vgLOCALIDAD = Mid(sLine, 148, 30)
         '      -------------------------------------------------------
                vCampo = "PROVINCIA"
                vPosicion = 178
        
                If Trim(Mid(sLine, 178, 2)) = "01" Then
                    vgPROVINCIA = "Capital Federal"
                ElseIf Trim(Mid(sLine, 178, 2)) = "02" Then
                    vgPROVINCIA = "Buenos Aires"
                ElseIf Trim(Mid(sLine, 178, 2)) = "03" Then
                    vgPROVINCIA = "Catamarca"
                ElseIf Trim(Mid(sLine, 178, 2)) = "04" Then
                    vgPROVINCIA = "Cordoba"
                ElseIf Trim(Mid(sLine, 178, 2)) = "05" Then
                    vgPROVINCIA = "Corrientes"
                ElseIf Trim(Mid(sLine, 178, 2)) = "06" Then
                    vgPROVINCIA = "Chaco"
                ElseIf Trim(Mid(sLine, 178, 2)) = "07" Then
                    vgPROVINCIA = "Chubut"
                ElseIf Trim(Mid(sLine, 178, 2)) = "08" Then
                    vgPROVINCIA = "Entre Rios"
                ElseIf Trim(Mid(sLine, 178, 2)) = "09" Then
                    vgPROVINCIA = "Formosa"
                ElseIf Trim(Mid(sLine, 178, 2)) = "10" Then
                    vgPROVINCIA = "Jujuy"
                ElseIf Trim(Mid(sLine, 178, 2)) = "11" Then
                    vgPROVINCIA = "La Pampa"
                ElseIf Trim(Mid(sLine, 178, 2)) = "12" Then
                    vgPROVINCIA = "La Rioja"
                ElseIf Trim(Mid(sLine, 178, 2)) = "13" Then
                    vgPROVINCIA = "Mendoza"
                ElseIf Trim(Mid(sLine, 178, 2)) = "14" Then
                    vgPROVINCIA = "Misiones"
                ElseIf Trim(Mid(sLine, 178, 2)) = "15" Then
                    vgPROVINCIA = "Neuquen"
                ElseIf Trim(Mid(sLine, 178, 2)) = "16" Then
                    vgPROVINCIA = "Rio Negro"
                ElseIf Trim(Mid(sLine, 178, 2)) = "17" Then
                    vgPROVINCIA = "Salta"
                ElseIf Trim(Mid(sLine, 178, 2)) = "18" Then
                    vgPROVINCIA = "San Juan"
                ElseIf Trim(Mid(sLine, 178, 2)) = "19" Then
                    vgPROVINCIA = "San Luis"
                ElseIf Trim(Mid(sLine, 178, 2)) = "20" Then
                    vgPROVINCIA = "Santa Cruz"
                ElseIf Trim(Mid(sLine, 178, 2)) = "21" Then
                    vgPROVINCIA = "Santa Fe"
                ElseIf Trim(Mid(sLine, 178, 2)) = "22" Then
                    vgPROVINCIA = "Santiago del Estero"
                ElseIf Trim(Mid(sLine, 178, 2)) = "23" Then
                    vgPROVINCIA = "Tierra del Fuego"
                ElseIf Trim(Mid(sLine, 178, 2)) = "24" Then
                    vgPROVINCIA = "Tucuman"
                End If
                '      -------------------------------------------------------
                vCampo = "TELEFONO"
                vPosicion = 180
                vgTelefono = Mid(sLine, 180, 15)
         '      -------------------------------------------------------
                vCampo = "MARCA"
                vPosicion = 195
                vgMARCADEVEHICULO = Mid(sLine, 195, 20)
         '      -------------------------------------------------------
                vCampo = "MODELO"
                vPosicion = 215
                vgMODELO = Mid(sLine, 215, 20)
         '      -------------------------------------------------------
                vCampo = "AÑO"
                vPosicion = 235
                vgAno = Mid(sLine, 235, 4)
         '      -------------------------------------------------------
         '         Atencion si la base es de Vehiculo va vgCOBERTURAVEHICULO
         '         Atencion si la base es de hogar va vgCOBERTURAHOGAR
                 vCampo = "CODIGO DE SERVICIO"
                vPosicion = 239
                vgCOBERTURAVEHICULO = Mid(sLine, 239, 2)
         '      -------------------------------------------------------
               vCampo = "FECHA INICIO COBERTURA"
                vPosicion = 241
                vgFECHAVIGENCIA = Mid(sLine, 241, 2) & "/" & Mid(sLine, 243, 2) & "/" & Mid(sLine, 245, 4)
         '      -------------------------------------------------------
                vCampo = "FECHA FIN COBERTURA"
                vPosicion = 250
                vgFECHAVENCIMIENTO = Mid(sLine, 249, 2) & "/" & Mid(sLine, 251, 2) & "/" & Mid(sLine, 253, 4)
         '      -------------------------------------------------------
                vCampo = "STATUS"
                vPosicion = 257
                vgOperacion = Mid(sLine, 257, 3)
         '      -------------------------------------------------------
                vCampo = "PRODUCTOR"
                vPosicion = 260
                vgCodigoDeProductor = Mid(sLine, 260, 10)
         '      -------------------------------------------------------
                vCampo = "Vip"
                vPosicion = 270
                vgCodigoDeServicioVip = Mid(sLine, 270, 1)
         '      -------------------------------------------------------
                 vCampo = "TIPODEVEHICULO"
                vPosicion = 271
                If Trim(Mid(sLine, 271, 20)) = "AUTOMOVIL" Then
                    vgTIPODEVEHICULO = 1
                ElseIf Trim(Mid(sLine, 271, 20)) = "4X4" Then
                    vgTIPODEVEHICULO = 3
                ElseIf Trim(Mid(sLine, 271, 20)) = "ACOPLADOS" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "AUTOMOVIL IMPORTADO" Then
                    vgTIPODEVEHICULO = 1
                ElseIf Trim(Mid(sLine, 271, 20)) = "CAMIONES > A.10 TN" Then
                    vgTIPODEVEHICULO = 4
                ElseIf Trim(Mid(sLine, 271, 20)) = "CAMIONES HASTA 10 TN" Then
                    vgTIPODEVEHICULO = 4
                ElseIf Trim(Mid(sLine, 271, 20)) = "CAMIONES HASTA 6 TN" Then
                    vgTIPODEVEHICULO = 4
                ElseIf Trim(Mid(sLine, 271, 20)) = "CASA RODANTE CON PROPULSION" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "CASA RODANTE SIN PROPULSION" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "COSECHADORAS FUMIGADORAS" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "IMPL. RURALES" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "MINIBUS" Then
                    vgTIPODEVEHICULO = 6
                ElseIf Trim(Mid(sLine, 271, 20)) = "MOTO  MENOR 50CC" Then
                    vgTIPODEVEHICULO = 5
                ElseIf Trim(Mid(sLine, 271, 20)) = "MOTO CHILENA" Then
                    vgTIPODEVEHICULO = 5
                ElseIf Trim(Mid(sLine, 271, 20)) = "MOTO MAYOR 50CC" Then
                    vgTIPODEVEHICULO = 5
                ElseIf Trim(Mid(sLine, 271, 20)) = "OMNIBUS" Then
                    vgTIPODEVEHICULO = 6
                ElseIf Trim(Mid(sLine, 271, 7)) = "PICK UP" And Trim(Mid(sLine, 281, 1)) = "A" Then
                    vgTIPODEVEHICULO = 3
                ElseIf Trim(Mid(sLine, 271, 7)) = "PICK UP" And Trim(Mid(sLine, 281, 1)) = "B" Then
                    vgTIPODEVEHICULO = 3
                ElseIf Trim(Mid(sLine, 271, 20)) = "SEMIRREMOLQUES" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "TRACTOR ACOPLADO" Then
                    vgTIPODEVEHICULO = 4
                ElseIf Trim(Mid(sLine, 271, 20)) = "TRACTORES" Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "TRAIL.BAT." Then
                    vgTIPODEVEHICULO = 7
                ElseIf Trim(Mid(sLine, 271, 20)) = "VAN" Then
                    vgTIPODEVEHICULO = 6
                ElseIf Trim(Mid(sLine, 271, 20)) = "VEHICULO CHILENO" Then
                    vgTIPODEVEHICULO = 7
                Else
                    vgTIPODEVEHICULO = 0
                End If
        ElseIf vidTipoDePoliza = 3 Then
            vTipoDePoliza = "Hogar"
         '      -------------------------------------------------------
                vCampo = "POLIZA"
                vPosicion = 1
                vgNROPOLIZA = Mid(sLine, 1, 15)
                
         '      -------------------------------------------------------
                vCampo = "APELLIDO Y NOMBRE CLIENTE"
                vPosicion = 21
                vgAPELLIDOYNOMBRE = Mid(sLine, 21, 50)
         '      -------------------------------------------------------
                vCampo = "TIPO DOCUMENTO"
                vPosicion = 71
                If Trim(Mid(sLine, 71, 2)) = "00" Then
                    vgTipodeDocumento = "CI Policia Federal"
                ElseIf Trim(Mid(sLine, 71, 2)) = "80" Then
                    vgTipodeDocumento = "CUIT"
                ElseIf Trim(Mid(sLine, 71, 2)) = "89" Then
                    vgTipodeDocumento = "LE"
                ElseIf Trim(Mid(sLine, 71, 2)) = "90" Then
                    vgTipodeDocumento = "LC"
                ElseIf Trim(Mid(sLine, 71, 2)) = "94" Then
                    vgTipodeDocumento = "Pasaporte"
                ElseIf Trim(Mid(sLine, 71, 2)) = "96" Then
                    vgTipodeDocumento = "DNI"
                ElseIf Trim(Mid(sLine, 71, 2)) = "99" Then
                    vgTipodeDocumento = "CUIT Externo"
                ElseIf Trim(Mid(sLine, 71, 2)) = "82" Then
                    vgTipodeDocumento = "CUIL"
                End If
         '      -------------------------------------------------------
                vCampo = "NUMERO DOCUMENTO"
                vPosicion = 73
                vgNumeroDeDocumento = Mid(sLine, 73, 11)
         '      -------------------------------------------------------
                vCampo = "DIRECCION"
                vPosicion = 84
                vgDOMICILIO = Mid(sLine, 84, 35)
         '      -------------------------------------------------------
                vCampo = "CODIGO POSTAL"
                vPosicion = 119
                vgCODIGOPOSTAL = Mid(sLine, 119, 5)
         '      -------------------------------------------------------
                vCampo = "LOCALIDAD"
                vPosicion = 124
                vgLOCALIDAD = Mid(sLine, 124, 30)
         '      -------------------------------------------------------
                vCampo = "PROVINCIA"
                vPosicion = 154
        
                If Trim(Mid(sLine, 154, 2)) = "01" Then
                    vgPROVINCIA = "Capital Federal"
                ElseIf Trim(Mid(sLine, 154, 2)) = "02" Then
                    vgPROVINCIA = "Buenos Aires"
                ElseIf Trim(Mid(sLine, 154, 2)) = "03" Then
                    vgPROVINCIA = "Catamarca"
                ElseIf Trim(Mid(sLine, 154, 2)) = "04" Then
                    vgPROVINCIA = "Cordoba"
                ElseIf Trim(Mid(sLine, 154, 2)) = "05" Then
                    vgPROVINCIA = "Corrientes"
                ElseIf Trim(Mid(sLine, 154, 2)) = "06" Then
                    vgPROVINCIA = "Chaco"
                ElseIf Trim(Mid(sLine, 154, 2)) = "07" Then
                    vgPROVINCIA = "Chubut"
                ElseIf Trim(Mid(sLine, 154, 2)) = "08" Then
                    vgPROVINCIA = "Entre Rios"
                ElseIf Trim(Mid(sLine, 154, 2)) = "09" Then
                    vgPROVINCIA = "Formosa"
                ElseIf Trim(Mid(sLine, 154, 2)) = "10" Then
                    vgPROVINCIA = "Jujuy"
                ElseIf Trim(Mid(sLine, 154, 2)) = "11" Then
                    vgPROVINCIA = "La Pampa"
                ElseIf Trim(Mid(sLine, 154, 2)) = "12" Then
                    vgPROVINCIA = "La Rioja"
                ElseIf Trim(Mid(sLine, 154, 2)) = "13" Then
                    vgPROVINCIA = "Mendoza"
                ElseIf Trim(Mid(sLine, 154, 2)) = "14" Then
                    vgPROVINCIA = "Misiones"
                ElseIf Trim(Mid(sLine, 154, 2)) = "15" Then
                    vgPROVINCIA = "Neuquen"
                ElseIf Trim(Mid(sLine, 154, 2)) = "16" Then
                    vgPROVINCIA = "Rio Negro"
                ElseIf Trim(Mid(sLine, 154, 2)) = "17" Then
                    vgPROVINCIA = "Salta"
                ElseIf Trim(Mid(sLine, 154, 2)) = "18" Then
                    vgPROVINCIA = "San Juan"
                ElseIf Trim(Mid(sLine, 154, 2)) = "19" Then
                    vgPROVINCIA = "San Luis"
                ElseIf Trim(Mid(sLine, 154, 2)) = "20" Then
                    vgPROVINCIA = "Santa Cruz"
                ElseIf Trim(Mid(sLine, 154, 2)) = "21" Then
                    vgPROVINCIA = "Santa Fe"
                ElseIf Trim(Mid(sLine, 154, 2)) = "22" Then
                    vgPROVINCIA = "Santiago del Estero"
                ElseIf Trim(Mid(sLine, 154, 2)) = "23" Then
                    vgPROVINCIA = "Tierra del Fuego"
                ElseIf Trim(Mid(sLine, 154, 2)) = "24" Then
                    vgPROVINCIA = "Tucuman"
                End If
                '      -------------------------------------------------------
                vCampo = "TELEFONO"
                vPosicion = 156
                vgTelefono = Mid(sLine, 156, 15)
         '      -------------------------------------------------------
         '         Atencion si la base es de Vehiculo va vgCOBERTURAVEHICULO
         '         Atencion si la base es de hogar va vgCOBERTURAHOGAR
                 vCampo = "CODIGO DE SERVICIO"
                vPosicion = 171
                vgCOBERTURAHOGAR = Mid(sLine, 171, 2)
         '      -------------------------------------------------------
               vCampo = "FECHA INICIO COBERTURA"
                vPosicion = 173
                vgFECHAVIGENCIA = Mid(sLine, 173, 2) & "/" & Mid(sLine, 175, 2) & "/" & Mid(sLine, 177, 4)
         '      -------------------------------------------------------
                vCampo = "FECHA FIN COBERTURA"
                vPosicion = 180
                vgFECHAVENCIMIENTO = Mid(sLine, 181, 2) & "/" & Mid(sLine, 183, 2) & "/" & Mid(sLine, 185, 4)
         '      -------------------------------------------------------
                vCampo = "STATUS"
                vPosicion = 189
                vgOperacion = Mid(sLine, 189, 3)
         '      -------------------------------------------------------
                vCampo = "PRODUCTOR"
                vPosicion = 192
                vgCodigoDeProductor = Mid(sLine, 192, 10)
         '      -------------------------------------------------------
                vCampo = "Vip"
                vPosicion = 202
                vgCodigoDeServicioVip = Mid(sLine, 202, 1)
         '      -------------------------------------------------------
        
        End If
                
                ssql = "Insert into bandejadeentrada.dbo.ImportaDatos347 ("
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
                ssql = ssql & "IdLote, IdTipodePoliza )"
                
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
                ssql = ssql & Trim(vidTipoDePoliza) & "') "
                cn.Execute ssql
                
                Ll = Ll + 1
                ll100 = ll100 + 1
                If ll100 = 100 Then
                    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " copiando linea " & Ll
                    ll100 = 0
                End If
                DoEvents
                '=======Blanqueavariables==============================
                vgPATENTE = ""
                vgNroMotor = ""
                vgNROPOLIZA = ""
                vgAPELLIDOYNOMBRE = ""
                vgTipodeDocumento = ""
                vgNumeroDeDocumento = ""
                vgDOMICILIO = ""
                vgCODIGOPOSTAL = ""
                vgLOCALIDAD = ""
                vgPROVINCIA = ""
                vgTelefono = ""
                vgMARCADEVEHICULO = ""
                vgMODELO = ""
                vgAno = ""
                vgCOBERTURAVEHICULO = ""
                vgCOBERTURAHOGAR = ""
                'vgFECHAVIGENCIA = ""
                'vgFECHAVENCIMIENTO = ""
                vgOperacion = ""
                vgCodigoDeProductor = ""
                vgCodigoDeServicioVip = ""
                vgTIPODEVEHICULO = 0
                '========================================================
    Loop
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & "  ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    vlineasTotales = Ll
    Ll = 0
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    
    cn1.Execute "update tm_campana set UltimaCorridaError='' , UltimaCorridaCantidadderegistros=0  where idcampana=" & vIDCampana
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
        ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros from tm_campana where idcampana=" & vIDCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        vRegistrosProcesados = vRegistrosProcesados + rsCMP("UltimaCorridaCantidaddeRegistros")
        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            lLote = vLote + 1 'para salir del FOR
        Else
                ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
        End If
        rsCMP.Close
    Next lLote
    cn1.Execute "update tm_campana set  UltimaCorridaCantidadderegistros = " & vRegistrosProcesados & " where idcampana=" & vIDCampana
Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub




