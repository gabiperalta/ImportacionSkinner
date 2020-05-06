Attribute VB_Name = "OrganizacionTorres"

Public Sub ImportarOrganizacionTorres()

Dim sssql As String, rsc As New Recordset, rsc2 As New Recordset
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
Dim vVigenciaVigente As String


On Error Resume Next
vgidCia = lIdCia ' sale del formulario del importador, al hacer click
vgidCampana = lIdCampana ' sale del formulario del importador, al hacer click


TablaTemporal

On Error Resume Next
 
Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application ' early binding el objeto excel
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
Set oSheet = oBook.Worksheets(1)
    
Dim filas As Long
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(5)
camposParaValidar(0) = "APELLIDO"
camposParaValidar(1) = "NOMBRE"
camposParaValidar(2) = "PATENTE"
camposParaValidar(3) = "POLIZA"
camposParaValidar(4) = "DNI"


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
'        rsc.AddNew
'====='Control de Lote===================================================
        nroLinea = nroLinea + 1
'        If lRow = 246 Then
'            MsgBox "Para"
'        End If
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
                Case "AÑO"
                    vgAno = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "COLOR"
                    vgCOLOR = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "APELLIDO"
                    vgAPELLIDOYNOMBRE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NOMBRE"
                    vgAPELLIDOYNOMBRE = vgAPELLIDOYNOMBRE & " " & v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PATENTE"
                    vgPATENTE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "POLIZA"
                    vgNROPOLIZA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MOTOR"
                    vgNroMotor = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MARCA Y MODELO"
                    vgMARCADEVEHICULO = Mid(v, 1, 50)
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DNI"
                    vgNumeroDeDocumento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DIRECCION"
                    vgDOMICILIO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "LOCALIDAD"
                    vgLOCALIDAD = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PROV"
                    vgPROVINCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "CP"
                    vgCODIGOPOSTAL = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TELEFONO"
                    vgTelefono = v
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
            
'        rsc.Close
        Next
        
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
            
'        vgFECHAVENCIMIENTO = DateAdd("m", 10, (rscn1("FECHAVENCIMIENTO")))
'        vVigenciaVigente = (rscn1("FECHAVIGENCIA"))
'        If Err.Number = 3021 Then 'Limpio error por registro Nuevo: El valor de BOF o EOF es True, o el actual registro se elimino
'            Err.Clear
'        End If
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                        If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
'                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vVigenciaVigente) Then vdif = vdif + 1
   '                    If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
   '                    If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                        If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                        If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                        If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                        vgIDPOLIZA = rscn1("idpoliza")
    '                   If vdif > 0 Then 'bloque para identificar modificaciones al hacer un debug.
    '                   vdif = vdif
    '
    '                     End If
                        
                    End If
                    
'                If vgIDPOLIZA = 0 Then
'                    vVigenciaVigente = Now
'                    vgFECHAVENCIMIENTO = DateAdd("m", 10, vVigenciaVigente)
'                End If
                
                rscn1.Close
    '=========='insert que se hace a la tabla temporal que se crea al comienzo==================
    
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
'           ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "ANO, "
            ssql = ssql & "COLOR, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "NroMotor, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
'            ssql = ssql & "FECHAVIGENCIA, "
'            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "CODIGOPOSTAL, "
            ssql = ssql & "Telefono, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", "
'           ssql = ssql & Trim(vgIdProducto) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgAno) & "', '"
            ssql = ssql & Trim(vgCOLOR) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
            ssql = ssql & Trim(vgNroMotor) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
'            ssql = ssql & Trim(vVigenciaVigente) & "', '"
'            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgTelefono) & "', "
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



Public Sub ImportarExelOrganizacionTorresOld()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim col As New Scripting.Dictionary
'Dim mExcel As New Excel.Application
'Dim wb
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

On Error GoTo errores

        ' Inicia Excel y abre el workbook
        Set oExcel = New Excel.Application
        oExcel.Visible = False
        Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
        Set oSheet = oBook.Worksheets(1)
'Dim sh As Excel.Sheets
    'Set mExcel = CreateObject("Excel.Application")
'    oExcel.Visible = False
'    Set oBooks = oExcel.Workbooks.Open(App.Path & "\" & sDirImportacion & "\" & FileImportacion, False, True)
        v = " "
'=======================================================
    vCtrolPatente = False
    'vCtrolVencimiento = False
    'vCtrolVigencia = False

'========================================================
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
'            Case "VIGHAS"
'                vCtrolVencimiento = True
'            Case "VIGDES"
'                vCtrolVigencia = True
            
        End Select

    Loop
    lCantCol = lCol

    If vCtrolPatente = False Then 'Or vCtrolVencimiento = False Or vCtrolVigencia = False Then
        MsgBox "Falta alguna Columna Obligatoria o esta mal la descripcion"
        Exit Sub
    End If

    If lCol = 1 Then
        MsgBox "Faltan campos"
        Exit Sub
    End If

    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatos", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 10000
        rsc.AddNew

        For lCol = 1 To lCantCol
            v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            If lCol = 1 And IsEmpty(v) Then Exit Do
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                Case "APELLIDO"
                    rsc("APELLIDOYNOMBRE").Value = v
'                Case "MARCA"
'                    rsc("MARCADEVEHICULO").Value = v
                Case "DESCRIPCION"
                    rsc("MODELO").Value = v
                Case "AÑO"
                    rsc("ANO").Value = v
'                Case "VIGDES"
'                    rsc("FECHAVIGENCIA").Value = v
'                Case "VIGHAS"
'                    rsc("FECHAVENCIMIENTO").Value = v
'                Case "LOCALIDAD"
'                Case "CP"
'                Case "RENUEVA"
'                Case ""
'                   rsc("IDPOLIZA") = v
'                Case ""
'                   rsc("IDCIA") = v
'                Case ""
'                   rsc("NUMEROCOMPANIA") = v
                Case "POLIZA"
                   rsc("NROPOLIZA") = v
'                Case ""
'                   rsc("NROSECUENCIAL") = v
                Case "NOMBRE"
                   rsc("APELLIDOYNOMBRE") = rsc("APELLIDOYNOMBRE") & ", " & v
'                Case "DOMICILIO"
'                   rsc("DOMICILIO") = v
                Case "LOCALIDAD"
                   rsc("LOCALIDAD") = v
                Case "PROV"
                   rsc("PROVINCIA") = v
                Case "CP"
                   rsc("CODIGOPOSTAL") = v
'                Case "VIG_DESDE"
'                   rsc("FECHAVIGENCIA") = v
'                Case "VIG_HASTA"
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
                Case "COLOR"
                   rsc("COLOR") = v
'                Case ""
'                   rsc("ANO") = v

'                Case ""
'                   rsc("PolizaIniciaAnualidad") = v
                Case "TELEFONO"
                   rsc("Telefono") = v
                Case "MOTOR"
                   rsc("NroMotor") = v
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
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    cn1.Execute sSPImportacion
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



