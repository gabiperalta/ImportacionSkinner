Attribute VB_Name = "GM"
Option Explicit
Public Sub ImportarGM()

Dim sssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim vExisteNroSecuencial As Integer
Dim i As Integer
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
    
'array para leer la primera row del excel y cargar los campos que trae el excel.
Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(9)
camposParaValidar(0) = "PATENTE"
camposParaValidar(1) = "NOMBRE"
camposParaValidar(2) = "MARCA"
camposParaValidar(3) = "MODELO"
camposParaValidar(4) = "ANIO"
camposParaValidar(5) = "VIGENCIA DESDE"
camposParaValidar(6) = "VIGENCIA HASTA"
camposParaValidar(7) = "LOCALIDAD"
camposParaValidar(8) = "CP"

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
    
    For posicion = 0 To 20
        vCoberturaLista(0, posicion) = "_"
        vCoberturaLista(1, posicion) = "_"
        vCoberturaLista(2, posicion) = "_"
    Next posicion

    Do While lCol < columnas + 1
        v = oSheet.Cells(1, lCol)
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
    Loop
    
    Do While lRow < filas + 1
    vCoberturaActual(0) = ""
    vCoberturaActual(1) = ""
    vCoberturaActual(2) = ""
    
    vExisteNroSecuencial = 0
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
                Case "PATENTE"
                    vgPATENTE = v
                    vgNROPOLIZA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NOMBRE"
                    vgAPELLIDOYNOMBRE = Replace(v, "'", "´")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MARCA"
                    vgMARCADEVEHICULO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MODELO"
                    vgMODELO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "ANIO"
                    vgAno = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "VIGENCIA DESDE"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "VIGENCIA HASTA"
                    vgFECHAVENCIMIENTO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "LOCALIDAD"
                    vgLOCALIDAD = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "CP"
                    vgCODIGOPOSTAL = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                Case "IDPRODUCTO"
'                    If Len(v) > 0 Then
'                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lIdCampana & "  and idproductoencliente = " & v
'                     rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly
'                        If Not rsprod.EOF Then
'                             vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
'                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                             vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
'                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                             vgCOBERTURAHOGAR = rsprod("coberturahogar")
'                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                             vgIdProducto = v
'                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                        Else
'                             vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", lRow, sName)
'
'                        End If
'                     rsprod.Close
'                    End If
                End Select
            End If
        Next
    
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
        ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("NROSECUENCIAL")) <> Trim(vgNROSECUENCIAL) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
'                        If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                        If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                        If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                        vgIDPOLIZA = rscn1("idpoliza")
                        
                    End If
        
                rscn1.Close
    '=========='insert que se hace a la tabla temporal que se crea al comienzo==================
    
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
            'ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "ANO, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "CODIGOPOSTAL, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "LOCALIDAD, "
            'ssql = ssql & "COBERTURAVEHICULO, "
            'ssql = ssql & "COBERTURAVIAJERO, "
            'ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
            'ssql = ssql & Trim(vgIdProducto) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgAno) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            'ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            'ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            'ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
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
            cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & lIdCia & ", " & lidCampana ' & ", " & vNombreTablaTemporal
            ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & lidCampana
            rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " Procesando los datos"
            ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
            DoEvents
            rsCMP.Close
        Next lLote
    
        cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & lIdCia & ", " & lidCampana

'============Finaliza Proceso========================================================
        cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lidCampana & ", " & vgCORRIDA
        CoberturasProcesadas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
        CoberturasBajas vgCORRIDA, vgidCampana, vgIdHistorialImportacion
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


Public Sub ImportarExelGMGrupoAsegurador()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax

Dim col As New Scripting.Dictionary
'Dim mExcel As New Excel.Application
'Dim wb

cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosGM"

On Error Resume Next
vgidCia = 10442
vgidCampana = 122

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet
' Inicia Excel y abre el workbook
Set oExcel = New Excel.Application
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
Set oSheet = oBook.Worksheets(1)
        
'======='control de lectura del archivo de datos
If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If
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
End If
'=======================================================================
    
Dim vctroldecol  As String
Dim vCtrolPatente As Boolean
Dim vCtrolNombre As Boolean
Dim vCtrolMarca As Boolean
Dim vCtrolModelo  As Boolean
Dim vCtrolAnio  As Boolean
Dim vCtrolVigenciaDesde  As Boolean
Dim vCtrolVigenciaHasta  As Boolean
Dim vCtrolCP  As Boolean
Dim vCtrolLocalidad  As Boolean
        
vCtrolPatente = False
vCtrolNombre = False
vCtrolMarca = False
vCtrolModelo = False
vCtrolAnio = False
vCtrolVigenciaDesde = False
vCtrolVigenciaHasta = False
vCtrolCP = False
vCtrolLocalidad = False
      
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
    
      
      
      lCol = 1
    lRow = 1
    Do While lCol < 50
        
        v = Trim(oSheet.Range(mToChar(lCol - 1) & "1").Value)
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
        Select Case v
            Case "PATENTE"
                vCtrolPatente = True
            Case "NOMBRE"
                vCtrolNombre = True
            Case "MARCA"
                vCtrolMarca = True
            Case "MODELO"
                vCtrolModelo = True
            Case "ANIO"
                vCtrolAnio = True
            Case "VIGENCIA DESDE"
                vCtrolVigenciaDesde = True
            Case "VIGENCIA HASTA"
                vCtrolVigenciaHasta = True
            Case "LOCALIDAD"
                vCtrolLocalidad = True
            Case "CP"
                vCtrolCP = True
            End Select
    Loop

    lCantCol = lCol
    vctroldecol = ""
    If vCtrolPatente = False Then vctroldecol = vctroldecol & ", " & "PATENTE"
    If vCtrolNombre = False Then vctroldecol = vctroldecol & ", " & "NOMBRE"
    If vCtrolMarca = False Then vctroldecol = vctroldecol & ", " & "MARCA"
    If vCtrolModelo = False Then vctroldecol = vctroldecol & ", " & "MODELO"
    If vCtrolAnio = False Then vctroldecol = vctroldecol & ", " & "ANIO"
    If vCtrolVigenciaDesde = False Then vctroldecol = vctroldecol & ", " & "VIGENCIA DESDE"
    If vCtrolVigenciaHasta = False Then vctroldecol = vctroldecol & ", " & "VIGENCIA HASTA"
    If vCtrolLocalidad = False Then vctroldecol = vctroldecol & ", " & "LOCALIDAD"
    If vCtrolCP = False Then vctroldecol = vctroldecol & ", " & "CP"
       
    If Len(vctroldecol) > 0 Then
        MsgBox "Faltan en el archivo las Siguientes Columnas: " & vctroldecol
        MsgBox "Importacion Detenida, Corrija el Archivo y vuelva a intentarlo!"
        Exit Sub
    End If
        

    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatosGM", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 10000
'=======Control de Lote===============================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'=====================================================
    
        rsc.AddNew
        rsc("idLOte") = vLote
        rsc("Modificaciones") = 1
        For lCol = 1 To lCantCol
            v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            If lCol = 1 And IsEmpty(v) Then Exit Do
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                    rsc("NROPOLIZA") = v
                Case "NOMBRE"
                    rsc("APELLIDOYNOMBRE").Value = v
                Case "MARCA"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "ANIO"
                    rsc("ANO").Value = v
                Case "VIGENCIA DESDE"
                    rsc("FECHAVIGENCIA").Value = v
                Case "VIGENCIA HASTA"
                    rsc("FECHAVENCIMIENTO").Value = v
                Case "CP"
                   rsc("CODIGOPOSTAL") = v
                Case "RENUEVA"
                Case "LOCALIDAD"
                   rsc("LOCALIDAD") = v
            End Select
        Next
        rsc.Update
'========Control de errores=========================================================
            If Err Then
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
                Err.Clear
            
            End If
'===========================================================================================
        lRow = lRow + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2
            ll100 = 0
        End If
        DoEvents

    Loop
'================Control de Leidos===============================================
            cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
            listoParaProcesar
'=================================================================================
    oExcel.Workbooks.Close
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
        ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " Procesando los datos"
        DoEvents
        rsCMP.Close
    Next lLote
    
    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
    
'============Finaliza Proceso========================================================
        cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lidCampana & ", " & vgCORRIDA
        Procesado
'=====================================================================================
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF

Exit Sub
errores:
    oExcel.Workbooks.Close
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow
    End If



End Sub



