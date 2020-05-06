Attribute VB_Name = "ColonSeguros"
Option Explicit

Public Sub ImportarColonMotos()

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

Dim camposParaValidar(4)
camposParaValidar(0) = "NROPOLIZA"
camposParaValidar(1) = "APELLIDOYNOMBRE"
camposParaValidar(2) = "DOCUMENTO"
camposParaValidar(3) = "INICIOVIGENCIA"
'camposParaValidar(4) = "IDPRODUCTO"


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
        
        'If lRow = 2848 Then
        '    MsgBox "Scocco"
        'End If

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
                Case "NROPOLIZA"
                    vgNROPOLIZA = Replace(v, "-", "")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "APELLIDOYNOMBRE"
                    vgAPELLIDOYNOMBRE = Replace(v, "'", "´")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPODEDOCUMENTO"
                    vgTipodeDocumento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DOCUMENTO"
                    vgNumeroDeDocumento = v
                    If IsEmpty(vgNROPOLIZA) Or Len(Trim(vgNROPOLIZA)) = 0 Then
                     vgNROPOLIZA = v
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                Case "VEHICULO"
'                    If UCase(Mid(v, 1, 4)) = "MOTO" Then
'                        v = "5"
'                    End If
'                    vgTIPODEVEHICULO = v
                Case "DOMINIO"
                    vgPATENTE = v
                    If Len(Trim(vgPATENTE)) > 13 Then
                        vgPATENTE = "  "
                    Else
                        vgPATENTE = v
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MARCA"
                    vgMARCADEVEHICULO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "MODELO"
                    vgMODELO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "IDPRODUCTO"
                    If Len(v) > 0 Then
                    ObtenerCoberturas lidCampana, (v), vCantDeErrores, flnErr, (lRow), (sName)
'                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente = " & v
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
                    End If
                Case "INICIOVIGENCIA"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FINVIGENCIA"
                    vgFECHAVENCIMIENTO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPODECLIENTE"
                    vgCodigoDeServicioVip = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                End Select
                vgTIPODEVEHICULO = "5"
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
         ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                        If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                        If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                        If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
   '                    If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
   '                    If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                        If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
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
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "Documento, "
            ssql = ssql & "TIPODEVEHICULO, "
            ssql = ssql & "PATENTE, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "MODELO, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "CodigoDeServicioVip, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", '"
            ssql = ssql & Trim(vgIdProducto) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgCodigoDeServicioVip) & "', "
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




Public Sub ImportarColonSeguros()


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
Dim vIdCampana As Long
Dim rsCMP As New Recordset
Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim vSinImportar As String
Dim vCoberturaFinanciera As String
Dim vFechaDeNacimiento As String
Dim vgFechaNacimiento As String
Dim vgVNumeroPoliza As String
Dim vFechaDeVigencia As String
Dim vFechaDeVencimiento As String

'Dim rs As ADODB.Recordset
'    rs.Open Null, Null, adOpenKeyset, adLockBatchOptimistic
'    rs.batchupdate

On Error GoTo errores
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosV2"
    vIDCIA = 10000803
   'vIDCampana = xxx
    LongDeLote = 1000

    Ll = 0
    sFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
    If Not fs.FileExists(sFile) Then Exit Sub
    Set tf = fs.OpenTextFile(sFile, ForReading, True)
    Ll = 1
    nroLinea = 1
    vLote = 1
    tf.SkipLine
    Do Until tf.AtEndOfStream
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
        
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
        
'Producto
        vCampo = "Producto"
        vPosicion = 1
        vgIdProducto = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Nro.Operación
        vCampo = "Nro.Operación"
        vPosicion = 2
        vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Id.Cliente
        vCampo = "Id.Cliente"
        vPosicion = 3
        vgCodigoEnCliente = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Apellido y Nombre/Razón Social
        vCampo = "Apellido y Nombre/Razón Social"
        vPosicion = 4
        vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Id.Tipo Doc
        vCampo = "Id.Tipo Doc"
        vPosicion = 5
        vgTipodeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Nro.Doc
        vCampo = "Nro.Doc"
        vPosicion = 6
        vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Fec.Nac
        vCampo = "Fec.Nac"
        vPosicion = 7
        vgFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Sexo
        vCampo = "Sexo"
        vPosicion = 8
        vgSexo = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Email
        vCampo = "Email"
        vPosicion = 9
        vgEmail = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Calle
        vCampo = "Calle"
        vPosicion = 10
        vgDOMICILIO = vgDOMICILIO & Trim(Mid(sLine, 1, InStr(1, sLine, ";") - 1))
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Altura
        vCampo = "Altura"
        vPosicion = 11
        vgDOMICILIO = vgDOMICILIO & Trim(Mid(sLine, 1, InStr(1, sLine, ";") - 1))
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Piso
        vCampo = "Piso"
        vPosicion = 12
        vgDOMICILIO = vgDOMICILIO & Trim(Mid(sLine, 1, InStr(1, sLine, ";") - 1))
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Dpto
        vCampo = "Dpto"
        vPosicion = 13
        vgDOMICILIO = vgDOMICILIO & Trim(Mid(sLine, 1, InStr(1, sLine, ";") - 1))
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Localidad
        vCampo = "Localidad"
        vPosicion = 14
        vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'CP
        vCampo = "CP"
        vPosicion = 15
        vgCODIGOPOSTAL = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Provincia
        vCampo = "Provincia"
        vPosicion = 16
        vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'País
        vCampo = "País"
        vPosicion = 17
        vgPais = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Teléfono1
        vCampo = "Teléfono1"
        vPosicion = 18
        vgTelefono = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Teléfono2
        vCampo = "Teléfono2"
        vPosicion = 19
        vgTelefono2 = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Teléfono3
        vCampo = "Teléfono3"
        vPosicion = 20
        vgTelefono3 = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Sponsor
        vCampo = "Sponsor"
        vPosicion = 21
        vgAgencia = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Id.Tip.Benef
        vCampo = "Id.Tip.Benef"
        vPosicion = 22
        vgDocumentoReferente = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Id.Tip.Asist
        vCampo = "Id.Tip.Asist"
        vPosicion = 23
        vgIdTipoDePOliza = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Id.Compañia
        vCampo = "Id.Compañia"
        vPosicion = 24
        vgidCia = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Adicional1
        vCampo = "Adicional1"
        vPosicion = 25
        vgMARCADEVEHICULO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Adicional2
        vCampo = "Adicional2"
        vPosicion = 26
        vgMODELO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Adicional3
        vCampo = "Adicional3"
        vPosicion = 27
        vgAno = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Adicional4
        vCampo = "Adicional4"
        vPosicion = 28
        vgTIPODEVEHICULO = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)
'Adicional5
        vCampo = "Adicional5"
        vPosicion = 29
        vgPATENTE = Mid(sLine, 1, InStr(1, sLine, ";") - 1)
        sLine = Mid(sLine, InStr(1, sLine, ";") + 1)

        If InStr(1, sLine, ";") > 0 Then
            MsgBox "Campos con ; que no corresponden en la linea" & Ll
            Exit Sub
        End If
        
        
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
Dim vcamp As Integer
Dim vdif As Long
Dim rscn1 As New Recordset
   ssql = "select *  from Auxiliout.dbo.tm_Polizas  where nroPoliza = '" & Trim(vgNROPOLIZA) & "' and Nrosecuencial = '" & vgNROSECUENCIAL & "'"
   rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
   
   vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
   vgIDPOLIZA = 0
           If Not rscn1.EOF Then
               vdif = 0  'setea la variale de control de repetido con modificacion en cero
               If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
               If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
               If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
               If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
               If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
               If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
               If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
               If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
               If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then vdif = vdif + 1
               If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
               If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
               If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
               If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
               If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
               If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
               If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
'                If Trim(rscn1("IDTIPODECOBERTURA")) <> Trim(vgIDTIPODECOBERTURA) Then vdif = vdif + 1
               If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
               If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
               If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
               If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then vdif = vdif + 1
               If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then vdif = vdif + 1
               If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then vdif = vdif + 1
'                If Trim(rscn1("IdCampana")) <> Trim(vgIdCampana) Then vdif = vdif + 1
               If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
               If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
               If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
               If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
               If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
               If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then vdif = vdif + 1
               If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then vdif = vdif + 1
               If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then vdif = vdif + 1
               If Trim(rscn1("FechNac")) <> Trim(vgFechaNacimiento) Then vdif = vdif + 1
               If Trim(rscn1("SEXO")) <> Trim(vgSexo) Then vdif = vdif + 1
               If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
               If Trim(rscn1("Telefono2")) <> Trim(vgTelefono2) Then vdif = vdif + 1
               If Trim(rscn1("Telefono3")) <> Trim(vgTelefono3) Then vdif = vdif + 1
               If Trim(rscn1("idproducto")) <> Trim(vgIdProducto) Then vdif = vdif + 1
               If Trim(rscn1("numeropoliza")) <> Trim(vgVNumeroPoliza) Then vdif = vdif + 1
               If Trim(rscn1("codigoencliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
               If Trim(rscn1("email")) <> Trim(vgEmail) Then vdif = vdif + 1
               If Trim(rscn1("Pais")) <> Trim(vgPais) Then vdif = vdif + 1
               If Trim(rscn1("Agencia")) <> Trim(vgAgencia) Then vdif = vdif + 1
               If Trim(rscn1("DocumentoReferente")) <> Trim(vgDocumentoReferente) Then vdif = vdif + 1
               If Trim(rscn1("IdTipoDePoliza")) <> Trim(vgIdTipoDePOliza) Then vdif = vdif + 1
               If Trim(rscn1("IdCia")) <> Trim(vgidCia) Then vdif = vdif + 1
               vgIDPOLIZA = rscn1("idpoliza")
           End If
                      
       rscn1.Close
'-=================================================================================================================
        
        
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatosV2 ("
        ssql = ssql & "IDPRODUCTO, "
        ssql = ssql & "NUMEROPOLIZA, "
        ssql = ssql & "CODIGOENCLIENTE, "
        ssql = ssql & "aPELLIDOYNOMBRE, "
        ssql = ssql & "TIPODEDOCUMENTO, "
        ssql = ssql & "NUMERODEDOCUMENTO, "
        ssql = ssql & "DOCUMENTO, "
        ssql = ssql & "FECHANAC, "
        ssql = ssql & "SEXO, "
        ssql = ssql & "EMAIL, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "CODIGOPOSTAL, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "PAIS, "
        ssql = ssql & "TELEFONO, "
        ssql = ssql & "TELEFONO2, "
        ssql = ssql & "TELEFONO3, "
        ssql = ssql & "AGENCIA, "
        ssql = ssql & "DOCUMENTOREFERENTE, "
        ssql = ssql & "IDTIPODEPOLIZA, "
        ssql = ssql & "IDCIA, "
        ssql = ssql & "MARCAVEHICULO, "
        ssql = ssql & "MODELO, "
        ssql = ssql & "ANO, "
        ssql = ssql & "TIPODEVEHICULO,"
        ssql = ssql & "PATENTE )"

        ssql = ssql & " values("
        ssql = ssql & Trim(vgIdProducto) & ", "
        ssql = ssql & Trim(vgNROPOLIZA) & ", '"
        ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgTipodeDocumento) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        ssql = ssql & Trim(vFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vgSexo) & "', '"
        ssql = ssql & Trim(vgEmail) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgLOCALIDAD) & "', '"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', "
        ssql = ssql & Trim(vgPROVINCIA) & ", '"
        ssql = ssql & Trim(vgPais) & "', '"
        ssql = ssql & Trim(vgTelefono) & "', '"
        ssql = ssql & Trim(vgTelefono2) & "', '"
        ssql = ssql & Trim(vgTelefono3) & "', '"
        ssql = ssql & Trim(vgAgencia) & "', "
        ssql = ssql & Trim(vgDocumentoReferente) & ", '"
        ssql = ssql & Trim(vgIdTipoDePOliza) & "', '"
        ssql = ssql & Trim(vgidCia) & "', '"
        ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
        ssql = ssql & Trim(vgMODELO) & "', '"
        ssql = ssql & Trim(vgAno) & "', '"
        ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
        ssql = ssql & Trim(vgPATENTE) & "') "

        cn.Execute ssql
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            ll100 = 0
        End If
        DoEvents
    Loop
    
    
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    vlineasTotales = Ll
    Ll = 0
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    'vUltimaCorrida As Long @nroCorrida as int
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
        ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vIDCIA & ", " & vIdCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vIdCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            lLote = vLote + 1 'para salir del FOR
        Else
            ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
            DoEvents
        End If
        rsCMP.Close
    Next lLote
    cn1.Execute "TM_BajaDePolizas" & " " & vUltimaCorrida & ", " & vIDCIA & ", " & vIdCampana
    
        

Exit Sub
errores:
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If


End Sub


