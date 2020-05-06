Attribute VB_Name = "BuenosAiresAsistencias"
Public Sub ImportarExelBuenosAiresAsistencias()

Dim sssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
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

Dim vCoberturaLista(0 To 2, 0 To 20) As String
Dim vLeidosPorCoberturaLista(0 To 2, 0 To 20) As Long
Dim vCoberturaActual(0 To 2) As String
Dim posicion As Integer

On Error Resume Next
vgidCia = lIdCia ' sale del formulario del importador, al hacer click
vgidCampana = lidCampana ' sale del formulario del importador, al hacer click


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
camposParaValidar(0) = "PATENTE"
camposParaValidar(1) = "Nº DE PÓLIZA"
camposParaValidar(2) = "FECHA DESDE"
camposParaValidar(3) = "FECHA HASTA"


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
    cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & vgidCampana & ", " & vgCORRIDA
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
    
    For posicion = 0 To 20
        vCoberturaLista(0, posicion) = "_"
        vCoberturaLista(1, posicion) = "_"
        vCoberturaLista(2, posicion) = "_"
    Next posicion
    
    Do While lRow < filas + 1
        vCoberturaActual(0) = ""
        vCoberturaActual(1) = ""
        vCoberturaActual(2) = ""
    
'====='Control de Lote===================================================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'===='Comienzo de lectura del excel======================================

        vCantDeErrores = 0
        For lCol = 1 To columnas
            sName = col.Item(lCol)
            v = oSheet.Cells(lRow, lCol)
            If IsEmpty(v) = False Then
            
            If lCol = 1 And IsEmpty(v) Then Exit Do
    
        Select Case UCase(Trim(sName))
                Case "PATENTE"
                    vgPATENTE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "Nº DE PÓLIZA"
                    vgNROPOLIZA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "APELLIDO Y NOMBRE"
                    vgAPELLIDOYNOMBRE = v
                    vgAPELLIDOYNOMBRE = Replace(vgAPELLIDOYNOMBRE, "'", "*")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "Marca del Vehículo"
                    vgMARCADEVEHICULO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MODELO"
                    vgMODELO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "ANO"
                    vgAno = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHA DESDE"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHA HASTA"
                    vgFECHAVENCIMIENTO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "TIPO DE SERVICIO"
                    vgTipodeServicio = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "TIPO DE VEHÍCULO"
                    vgIDTIPODECOBERTURA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "COBERTURA VEHÍCULO"
                    vgCOBERTURAVEHICULO = v
                    If Len(vgCOBERTURAVEHICULO) < 2 Then
                        vgCOBERTURAVEHICULO = "0" + vgCOBERTURAVEHICULO
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "COBERTURA VIAJERO"
                    vgCOBERTURAVIAJERO = v
                    If Len(vgCOBERTURAVIAJERO) < 2 Then
                        vgCOBERTURAVIAJERO = "0" + vgCOBERTURAVIAJERO
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "DOMICILIO"
                    vgDOMICILIO = Replace(v, "'", "")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName) ' cuenta el error en el campo, si lo hubiere.
                Case "Provincia"
                   vgPROVINCIA = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "Localidad"
                   vgLOCALIDAD = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "VIP"
                    vgCodigoDeServicioVip = v
                
                End Select
            End If
        Next
        
    If vCantDeErrores = 0 Then
    
        '=============== Lectura de coberturas ===============
        vCoberturaActual(0) = vgCOBERTURAVEHICULO
        vCoberturaActual(1) = vgCOBERTURAVIAJERO
        vCoberturaActual(2) = vgCOBERTURAHOGAR
        
        Dim coberturaPosicion As Integer
        
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
        Next coberturaPosicion
        
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
         ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                        If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                        If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
                        If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                        If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                        If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                        If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                        vgIDPOLIZA = rscn1("idpoliza")
    '                   If vdif > 0 Then 'bloque para identificar modificaciones al hacer un debug.
    '                   vdif = vdif
    '
    '                     End If
                        
                    End If
        
                rscn1.Close
    '=========='insert que se hace a la tabla temporal que se crea al comienzo==================
    
            ssql = "Insert into BandejaDeEntrada.dbo.importaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "ANO, "
            ssql = ssql & "TipodeServicio, "
            ssql = ssql & "IDTIPODECOBERTURA, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "CodigoDeServicioVip, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgAno) & "', '"
            ssql = ssql & Trim(vgTipodeServicio) & "', '"
            ssql = ssql & Trim(vgIDTIPODECOBERTURA) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", '"
            ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
            ssql = ssql & Trim(vLote) & "', '"
            ssql = ssql & Trim(vdif) & "') "
            cn.Execute ssql
            
            
    '========Control de errores=========================================================
    
                If Err Then
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
                    Err.Clear
                
                End If
                    
    '===========================================================================================
    End If

' bloque para el control de los registros leidos que se plasma (4K-HD) en la App del importador.
        If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        lRow = lRow + 1
        ll100 = ll100 + 1
        
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow

'========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lidCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql

            ll100 = 0
        End If
        DoEvents
    Loop
    
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
            cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
            ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
            rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " Procesando los datos"
            ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
            DoEvents
            rsCMP.Close
        Next lLote
    
        cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso========================================================
        cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & vgidCampana & ", " & vgCORRIDA
        CoberturasProcesadas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
        CoberturasBajas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
        Procesado
'====================================================================================
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF


Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

oExcel.Workbooks.Close
Set oExcel = Nothing

End Sub





Public Sub ImportarExelBuenosAiresAsistenciasOld()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vCtrolPatente As Boolean, vCtrolVigencia As Boolean, vCtrolVencimiento As Boolean
Dim col As New Scripting.Dictionary
'Dim mExcel As New Excel.Application
'Dim wb
'        Dim oExcel As Excel.Application
'        Dim oBook As Excel.Workbook
'        Dim oSheet As Excel.Worksheet

On Error GoTo errores

        ' Inicia Excel y abre el workbook
'        Set oExcel = New Excel.Application
'        oExcel.Visible = False
'        Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
'        Set oSheet = oBook.Worksheets(1)
'Dim sh As Excel.Sheets
    'Set mExcel = CreateObject("Excel.Application")
'    oExcel.Visible = False
'    Set oBooks = oExcel.Workbooks.Open(App.Path & "\" & sDirImportacion & "\" & FileImportacion, False, True)
' Inicia Excel y abre el workbook

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
'Set oBooks = oExcel.Workbooks
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
'oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Sheets(1)
        
        '=====inicio del control de corrida====================================
                        vgCORRIDA = ControlDeCorrida()
        '======================================================================
        
        
        
    v = " "
    vCtrolPatente = False
    vCtrolVencimiento = False
    vCtrolVigencia = False

    lCol = 1
    lRow = 1
    Do While lCol < 50
        v = oSheet.Range(mToChar(lCol - 1) & "1").Value
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
        Select Case v
            Case "PATENTE"
                vCtrolPatente = True
            Case "VIGHAS"
                vCtrolVencimiento = True
            Case "VIGDES"
                vCtrolVigencia = True
            
        End Select
    Loop
    lCantCol = lCol

'    If vCtrolPatente = False Or vCtrolVencimiento = False Or vCtrolVigencia = False Then
'        MsgBox "Falta alguna Columna Obligatoria o esta mal la descripcion"
'        Exit Sub
'    End If

    If lCol = 1 Then
        MsgBox "Faltan campos"
        Exit Sub
    End If

    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosBuenosAiresAsistencias"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatosBuenosAiresAsistencias", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 30000
        rsc.AddNew

        For lCol = 1 To lCantCol
            'v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            v = oSheet.Cells(lRow, lCol)
            If lCol = 1 And IsEmpty(v) Then Exit Do
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                Case "Nº DE PÓLIZA"
                    rsc("NROPOLIZA").Value = v
                Case "APELLIDO Y NOMBRE"
                    rsc("APELLIDOYNOMBRE").Value = v
                Case "Marca del Vehículo"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "ANO"
                    rsc("ANO").Value = v
                Case "FECHA DESDE"
                    rsc("FECHAVIGENCIA").Value = v
                Case "FECHA HASTA"
                    rsc("FECHAVENCIMIENTO").Value = v
'                Case "LOCALIDAD"
'                Case "CP"
'                Case "RENUEVA"
'                Case ""
'                   rsc("IDPOLIZA") = v
'                Case ""
'                   rsc("IDCIA") = v
'                Case ""
'                   rsc("NUMEROCOMPANIA") = v
'                Case ""
'                Case ""
'                   rsc("NROSECUENCIAL") = v
'                Case ""
'                   rsc("APELLIDOYNOMBRE") = v
                Case "DOMICILIO"
                   rsc("DOMICILIO") = v
                Case "Localidad"
                   rsc("LOCALIDAD") = v
                Case "Provincia"
                   rsc("PROVINCIA") = v
'                Case ""
'                   rsc("CODIGOPOSTAL") = v
'                Case ""
'                   rsc("FECHAVIGENCIA") = v
'                Case ""
'                   rsc("FECHAVENCIMIENTO") = v
'                Case ""
'                   rsc("FECHAALTAOMNIA") = v
'                Case ""
'                   rsc("FECHABAJAOMNIA") = v
'                Case ""
'                   rsc("IDAUTO") = v
'                Case ""
'                   rsc("MARCADEVEHICULO") = v
'                Case ""
'                   rsc("MODELO") = v
'                Case ""
'                   rsc("COLOR") = v
'                Case ""
'                   rsc("ANO") = v
'                Case ""
'                   rsc("PATENTE") = v
'                Case ""
'                   rsc("TIPODEVEHICULO") = v
                Case "TIPO DE SERVICIO"
                   rsc("TipodeServicio") = v
                Case "TIPO DE VEHÍCULO"
                   rsc("IDTIPODECOBERTURA") = v
                Case "COBERTURA VEHÍCULO"
                   rsc("COBERTURAVEHICULO") = v
                Case "COBERTURA VIAJERO"
                   rsc("COBERTURAVIAJERO") = v
'                Case ""
'                   rsc("TipodeOperacion") = v
'                Case ""
'                   rsc("Operacion") = v
'                Case ""
'                   rsc("CATEGORIA") = v
'                Case ""
'                   rsc("ASISTENCIAXENFERMEDAD") = v
'                Case ""
'                   rsc("CORRIDA") = v
'                Case ""
'                   rsc("FECHACORRIDA") = v
'                Case ""
'                   rsc("IdCampana") = v
'                Case ""
'                   rsc("Conductor") = v
'                Case ""
'                   rsc("CodigoDeProductor") = v
'                Case ""
'                   rsc("CodigoDeServicioVip") = v
'                Case ""
'                   rsc("TipodeDocumento") = v
'                Case ""
'                   rsc("NumeroDeDocumento") = v
'                Case ""
'                   rsc("TipodeHogar") = v
'                Case ""
'                   rsc("IniciodeAnualidad") = v
'                Case ""
'                   rsc("PolizaIniciaAnualidad") = v
'                Case ""
'                   rsc("Telefono") = v
'                Case ""
'                   rsc("NroMotor") = v
'                Case ""
'                   rsc("Gama") = v
'                Case ""
'                   rsc("NroDocumento") = v
            End Select
        Next
        rsc.Update
        lRow = lRow + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2
            ll100 = 0
        End If
        DoEvents

    Loop
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    cn1.Execute sSPImportacion
Exit Sub
errores:
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow
    End If



End Sub


