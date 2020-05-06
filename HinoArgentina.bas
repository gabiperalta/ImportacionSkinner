Attribute VB_Name = "HinoArgentina"
Public Sub ImportarHino()

Dim sssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vIDCampana As Long
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim vlineasTotales As Long
Dim sArchivo As String
Dim ssql As String
Dim ssssql As String
Dim rsprod As New Recordset
Dim regMod As Long
Dim vNombreTablaTemporal As String


On Error Resume Next
vgidCia = lIdCia ' sale del formulario del importador, al hacer click
vgidCampana = lIdCampana ' sale del formulario del importador, al hacer click


TablaTemporal ' procedimiento que crea la tabla temporal de manera dinamica toma el valor del idcampana y lo concatena al nombre de la tabla temporal .


On Error Resume Next
 
Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application ' early binding el objeto excel
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
Set oSheet = oBook.Worksheets(1)
    
Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(3)
camposParaValidar(0) = "NROPOLIZA"
camposParaValidar(1) = "APELLIDOYNOMBRE"
'camposParaValidar(2) = "DOCUMENTO"
'camposParaValidar(2) = "INICIOVIGENCIA"
camposParaValidar(2) = "IDPRODUCTO"


'========'objeto excel para almacenar errores============================

If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    Dim vCantDeErrores As Integer
    Dim sFileErr As New FileSystemObject
    Dim flnErr As TextStream
    Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
    flnErr.WriteLine "Errores"
    vCantDeErrores = 0

'======='control de lectura del archivo de datos=========================
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Exit Sub
    End If
'======'inicio del control de corrida====================================
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
'======'incializacion de variables de control de lote====================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
    lRow = 2
    lCol = 1

    Do While lCol < columnas + 1
        v = oSheet.Cells(1, lCol)
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
    Loop
    
    Do While lRow < filas + 1

'====='Control de Lote===================================================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'===='Comienzo de lectura del excel======================================
        Blanquear

        vCantDeErrores = 0
        For lCol = 1 To columnas
            sName = col.Item(lCol)
            v = oSheet.Cells(lRow, lCol)
            If IsEmpty(v) = False Then
            
            If lCol = 1 And IsEmpty(v) Then Exit Do
    
        Select Case UCase(Trim(sName))
                Case "APELLIDOYNOMBRE"
                    vgAPELLIDOYNOMBRE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "DOCUMENTO"
                    vgNumeroDeDocumento = v
                    If IsEmpty(vgNROPOLIZA) Or Len(Trim(vgNROPOLIZA)) = 0 Then
                     vgNROPOLIZA = v
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPO DE DOCUMENTO"
                    vgTipodeDocumento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "PATENTE"
                    vgPATENTE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "INICIOVIGENCIA"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FINVIGENCIA"
                    vgFECHAVENCIMIENTO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHANACIMIENTO"
                    vgFechaDeNacimiento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "IDPRODUCTO"
                
                    If Len(v) > 0 Then
                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lIdCampana & "  and idproductoencliente = " & v
                     rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly
                        If Not rsprod.EOF Then
                             vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgCOBERTURAHOGAR = rsprod("coberturahogar")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgIdProducto = v
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                        Else
                             vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", lRow, sName)
                        
                        End If
                     rsprod.Close
                    End If
                Case "DOMICILIO"
                   vgDOMICILIO = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "CODIGOPOSTAL"
                   vgCODIGOPOSTAL = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MODELO"
                    vgMODELO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PROVINCIA"
                   vgPROVINCIA = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "LOCALIDAD"
                   vgLOCALIDAD = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TELEFONO"
                   vgTelefono = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NROPOLIZA"
                    vgNroMotor = v
                    vgNROPOLIZA = v
                    vgNROPOLIZA = Right(v, 7) 'UCase(Mid(v, 12, 6))
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)

                End Select
                
            vgTIPODEVEHICULO = "4"
            vgMARCADEVEHICULO = "HINO"
            End If
        Next
        
        If vgFECHAVENCIMIENTO = "00:00:00" Then
            vgFECHAVENCIMIENTO = DateAdd("yyyy", 3, vgFECHAVIGENCIA)
        End If
        
'    If vCantDeErrores = 0 Then
    '==============  IMPORTANTE   ================================================================.
    ' bloque de verificacion de modificaciones.
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
    
         Dim vcamp As Integer
         Dim vdif As Long
         Dim rscn1 As New Recordset
         ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lIdCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("TIPO DE DOCUMENTO")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("CodigoEnCliente"))) <> Trim(vgIdProducto) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                        If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                        If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                        If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                        If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
    '                   If Trim(rscn1("idcia")) <> Trim(lIdCia) Then vdif = vdif + 1
                        vgIDPOLIZA = rscn1("idpoliza")
    '                   If vdif > 0 Then 'bloque para identificar modificaciones al hacer un debug.
    '                   vdif = vdif
    '
    '                     End If
                        
                    End If
        
                rscn1.Close
    '=========='insert que se hace a la tabla temporal que se crea al comienzo==================
    
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
            ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "NroMotor, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "FechadeNacimiento, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "CODIGOPOSTAL, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "Telefono, "
            ssql = ssql & "TIPODEVEHICULO, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", '"
            ssql = ssql & Trim(vgIdProducto) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgNroMotor) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgTelefono) & "', '"
            ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", '"
            ssql = ssql & Trim(vLote) & "', '"
            ssql = ssql & Trim(vdif) & "') "
            cn.Execute ssql
            
            
    '========Control de errores=========================================================
    
                If Err Then
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
                    Err.Clear
                
                End If
                    
    '===========================================================================================
' bloque para el control de los registros leidos que se plasma (4K-HD) en la App del importador.
        If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        lRow = lRow + 1
        ll100 = ll100 + 1
        
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow

'========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lIdCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql

            ll100 = 0
        End If
        DoEvents
    Loop

'================Control de Leidos===========llama al storeprocedure para hacer un update en tm_importacionHistorial

        cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
        listoParaProcesar
                            
'========================================plasma en el cuadro de mensaje de la app los valores del conteo de linea por lote
'y genera cuadro de mensaje para procesar los registros leidos.
    
        ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
        If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
'===============inicio del Control de Procesos=====================================

        cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
                            
'==================================================================================
        ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
        Dim rsCMP As New Recordset
        DoEvents
        For lLote = 1 To vLote
            cn1.CommandTimeout = 300
            cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & lIdCia & ", " & lIdCampana ' & ", " & vNombreTablaTemporal
            ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & lIdCampana
            rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " Procesando los datos"
            ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
            DoEvents
            rsCMP.Close
        Next lLote
    
        cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & lIdCia & ", " & lIdCampana

'============Finaliza Proceso========================================================
        cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
        Procesado
'====================================================================================
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF

'  cn1.Execute "TM_BajaDePolizas" & " " & vgCORRIDA & ", " & vgidcia & ", " & vgidcampana
Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

oExcel.Workbooks.Close
Set oExcel = Nothing

End Sub
