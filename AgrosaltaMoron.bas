Attribute VB_Name = "AgrosaltaMoron"
Public Sub ImportarExelAgrosaltaMoron()

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
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
Set oSheet = oBook.Worksheets(1)
    
'array para leer la primera row del excel y cargar los campos que trae el excel.
Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(3)
camposParaValidar(0) = "PATENTE"
camposParaValidar(1) = "VIGHAS"
camposParaValidar(2) = "VIGDES"

'========'objeto excel para almacenar errores============================

If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    Dim vCantDeErrores As Integer
    Dim sFileErr As New FileSystemObject
    Dim flnErr As TextStream
    Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
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
                    vgPATENTE = Replace(v, "-", "")
                    vgNROPOLIZA = vgPATENTE
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
                Case "VIGDES"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "VIGHAS"
                    vgFECHAVENCIMIENTO = v
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
        ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lIdCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
'                        If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
'                        If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
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
            ssql = ssql & "MODELO, "
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
            ssql = ssql & Trim(vgMODELO) & "', '"
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


Public Sub ImportarExelAgrosaltaMoronOld()
Dim ssql As String, rsc As New Recordset, rsCorr As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vCtrolPatente As Boolean, vCtrolVigencia As Boolean, vCtrolVencimiento As Boolean
Dim col As New Scripting.Dictionary
Dim lLote As Long
Dim vLote As Long
Dim nroLinea As Long
'Dim mExcel As New Excel.Application
'Dim wb
'        Dim oExcel As Excel.Application
'        Dim oBook As Excel.Workbook
'        Dim oSheet As Excel.Worksheet

On Error Resume Next
vLote = 1
LongDeLote = 1000
nroLinea = 0

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
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
        Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
        'oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets(1)
        
        'control de lectura del archivo de datos
        If Err Then
            MsgBox Err.Description
            Err.Clear
            Exit Sub
        End If
        
'control de campos====================================
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

    If vCtrolPatente = False Or vCtrolVencimiento = False Or vCtrolVigencia = False Then
        MsgBox "Falta alguna Columna Obligatoria o esta mal la descripcion"
        Exit Sub
    End If

    If lCol = 1 Then
        MsgBox "Faltan campos"
        Exit Sub
    End If
'=====================================================================================

    cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lIdCampana & ", " & vgCORRIDA
    
    ssql = "Select top 1 corrida from tm_ImportacionHistorial where idcampana = " & lIdCampana & " and Registrosleidos is null"
    rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    If rsCorr.EOF Then
        MsgBox "no se determino la corrida, se detendra el proceso"
        Exit Sub
    Else
        vgCORRIDA = rsCorr("corrida")
    End If
    
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosagrosaltaMoron"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatosagrosaltaMoron", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 30000
        rsc.AddNew
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
        rsc("idlote") = vLote
        

        For lCol = 1 To lCantCol
            'v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            v = oSheet.Cells(lRow, lCol)
            
            
        

            If lCol = 1 And IsEmpty(v) Then Exit Do
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                    'v.Replace("-", "")
                    rsc("NROPOLIZA").Value = v
                Case "NOMBRE"
                    rsc("APELLIDOYNOMBRE").Value = v
                Case "MARCA"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "ANIO"
                    rsc("ANO").Value = v
                Case "VIGDES"
                    rsc("FECHAVIGENCIA").Value = v
                Case "VIGHAS"
                    rsc("FECHAVENCIMIENTO").Value = v
                Case "LOCALIDAD"
                Case "CP"
                Case "RENUEVA"

            End Select
        Next
        rsc.Update
            If Err Then
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
                Err.Clear
            
            End If

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
    
        If Err Then
         MsgBox "Error " & Err.Description
         Err.Clear
        End If
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
    
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
         ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
   
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & vgCORRIDA & "," & lLote '& ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIDCampana
'        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vIDCampana
'        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If Err Then
         MsgBox "Error " & Err.Description
         Err.Clear
'         Exit Sub
        End If
'        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
'            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
'            lLote = vLote + 1 'para salir del FOR
'        Else
                ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
'        End If
'        rsCMP.Close
    Next lLote
cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF
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

