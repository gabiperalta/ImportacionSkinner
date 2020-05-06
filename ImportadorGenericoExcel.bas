Attribute VB_Name = "ImportadorGenericoExcel"
Option Explicit

Public Sub ImportarGenericoExcel()

Dim sssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim vExisteNroSecuencial As Integer
Dim vInformaFecha As Integer
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
Dim vVigenciaVigente As Date
Dim fechaActual As Date

fechaActual = Now

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
Dim filas As Long
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(1)
'camposParaValidar(0) = "NROPOLIZA"
'camposParaValidar(1) = "APELLIDOYNOMBRE"
'camposParaValidar(3) = "DOCUMENTO"
'camposParaValidar(4) = "INICIOVIGENCIA"
camposParaValidar(1) = "IDPRODUCTO"


'========'objeto excel para almacenar errores============================

If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    Dim vCantDeErrores As Integer
    Dim sFileErr As New FileSystemObject
    Dim flnErr As TextStream
    Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
    'Set flnErr = sFileErr.CreateTextFile(appPathTemp & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
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
    vExisteNroSecuencial = 0
    vInformaFecha = 0
'====='Control de Lote===================================================
        nroLinea = nroLinea + 1
'        If lRow = 265 Then
'            MsgBox "para"
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
                Case "APELLIDOYNOMBRE"
                    vgAPELLIDOYNOMBRE = Replace(v, "'", "")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NOMBRE"
                    vgNombre = v
                    vgAPELLIDOYNOMBRE = vgAPELLIDOYNOMBRE & " " & vgNombre
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "APELLIDO"
                    vgApellido = v
                    vgAPELLIDOYNOMBRE = vgApellido
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NROSECUENCIAL"
                    If Not IsEmpty(vgNROSECUENCIAL) Then
                        vgNROSECUENCIAL = v
                        vExisteNroSecuencial = 1
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DOCUMENTO"
                    vgNumeroDeDocumento = v
                    If IsEmpty(vgNROPOLIZA) Or Len(Trim(vgNROPOLIZA)) = 0 Then
                        vgNROPOLIZA = Replace(v, "-", "")
                        vgNROPOLIZA = Replace(vgNROPOLIZA, ".", "")
                        If IsNumeric(vgNROPOLIZA) Then vgNROPOLIZA = CStr(CLng(vgNROPOLIZA))
                    End If
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "CARGO"
                    vgCargo = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPODEDOCUMENTO"
                    vgTipodeDocumento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PATENTE"
                    vgPATENTE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "INICIOVIGENCIA"
                    vVigenciaVigente = v
                    'vFECHAVIGENCIA = v
                    If vVigenciaVigente = "00:00:00" Then
                        vVigenciaVigente = Now
                    End If
                    'vFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FINVIGENCIA"
                    vgFECHAVENCIMIENTO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHANACIMIENTO"
                    'vgFechaDeNacimiento = v
                    If Not IsDate(v) Then
                        vgFechaDeNacimiento = "00:00:00"
                        'vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                    Else
                        vgFechaDeNacimiento = v
                    End If
                    
                Case "IDPRODUCTO"
                    If Len(v) > 0 Then
                    ObtenerCoberturas lidCampana, (v), vCantDeErrores, flnErr, (lRow), (sName)
'                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente = '" & v & "'"
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
                Case "DOMICILIO"
                   vgDOMICILIO = Replace(v, "'", "")
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
                   vgLOCALIDAD = Replace(v, "'", "")
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TELEFONO"
                   vgTelefono = Mid(v, 1, 20)
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NROPOLIZA"
                    vgNROPOLIZA = Replace(v, "-", "")
                    vgNROPOLIZA = Replace(vgNROPOLIZA, ".", "")
                    If IsNumeric(vgNROPOLIZA) Then vgNROPOLIZA = CStr(CLng(vgNROPOLIZA))
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                End Select
            End If
        Next
    
    
    '1004 es ASISTENCIA GLOBAL | 782 es AMMMA | 988 es EMPLEADOS COLON | 1005 es LA RED DE SERVICIOS
    '978 es ARCOR | 62 es ASMEPRIV | 431 es AMEP | 708 es TARJETAACTUAL
    
    If vgidCampana = 782 Or vgidCampana = 988 Or vgidCampana = 1005 Or vgidCampana = 978 Or vgidCampana = 431 Or vgidCampana = 708 Or vgidCampana = 1127 Then
        v = DateAdd("yyyy", 1, fechaActual) 'hasta el 10/10/19 era vVigenciaVigente
        vgFECHAVENCIMIENTO = v
    End If
    
    If vgFECHAVENCIMIENTO = "00:00:00" Then 'si no vino la fecha de vencimiento en la base y falta menos de dos meses para que venza, agregame un anio plis
        If DateDiff("m", vgFECHAVENCIMIENTO, fechaActual) < 2 Then
            v = DateAdd("yyyy", 1, vgFECHAVENCIMIENTO)
            vgFECHAVENCIMIENTO = v
        End If
    End If
    
    If vgidCampana = 62 Then 'para Hino se corta el numero de poliza
        v = Mid(vgNROPOLIZA, 4)
        vgNROPOLIZA = v
    End If
    
    '==============  IMPORTANTE   ================================================================'
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
         If vExisteNroSecuencial = 1 Then
            ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' and Nrosecuencial = '" & vgNROSECUENCIAL & "'"
         Else
            ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
         End If
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

        'vgFECHAVENCIMIENTO = DateAdd("m", 2, Now)
        vVigenciaVigente = (rscn1("FECHAVIGENCIA"))
        If Err.Number = 3021 Then 'Limpio error: El valor de BOF o EOF es True, o el actual registro se elimino. Se genera porque el registro llego por primera vez
            Err.Clear
        End If
            
        vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
        vgIDPOLIZA = 0
                If Not rscn1.EOF Then
                    vdif = 0  'setea la variale de control de repetido con modificacion en cero
                    If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                    If Trim(rscn1("NROSECUENCIAL")) <> Trim(vgNROSECUENCIAL) Then vdif = vdif + 1
                    If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                    If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                    If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                    If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                    If vgFechaDeNacimiento Then
                        If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                    End If
                    If vVigenciaVigente Then
                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vVigenciaVigente) Then vdif = vdif + 1
                    End If
                    If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                    If vgFECHAVENCIMIENTO Then
                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                    End If
                    If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
                    If vgCOBERTURAVEHICULO <> "" Then
                        If CInt(Trim(rscn1("COBERTURAVEHICULO"))) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                    End If
                    If vgCOBERTURAVIAJERO <> "" Then
                        If CInt(Trim(rscn1("COBERTURAVIAJERO"))) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                    End If
                    If vgCOBERTURAHOGAR <> "" Then
                        If CInt(Trim(rscn1("COBERTURAHOGAR"))) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                    End If
                    If vgCOBERTURAAP <> "" Then
                        If CInt(Trim(rscn1("COBERTURAAP"))) <> Trim(vgCOBERTURAAP) Then vdif = vdif + 1
                    End If
                    If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                    If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                    If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                    If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                    If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                    If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                    If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                    If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                    vgIDPOLIZA = rscn1("idpoliza")
                    
                End If
                
                If vgFECHAVENCIMIENTO = "00:00:00" Then
                
                    If vgIDPOLIZA = 0 Then
                        vVigenciaVigente = fechaActual
                        vgFECHAVENCIMIENTO = DateAdd("m", 6, vVigenciaVigente) ' hasta el 7/10/2019 eran 2 meses
                    ElseIf vgIDPOLIZA <> 0 Then
                        vgFECHAVENCIMIENTO = DateAdd("m", 6, fechaActual) ' hasta el 7/10/2019 eran 2 meses
                        vdif = vdif + 1
                    End If
                
                End If
                                
            rscn1.Close
    '=========='insert que se hace a la tabla temporal que se crea al comienzo==================
    
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
            ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "NROSECUENCIAL, "
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
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "CARGO, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "COBERTURAAP, "
            ssql = ssql & "Telefono, "
            ssql = ssql & "TIPODEVEHICULO, "
            ssql = ssql & "MARCADEVEHICULO, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "ModificacionesEnFechas,"
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", '"
            ssql = ssql & Trim(vgIdProducto) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
            ssql = ssql & Trim(vgPATENTE) & "', '"
            ssql = ssql & Trim(vVigenciaVigente) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
            ssql = ssql & Trim(vgMODELO) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgCargo) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgCOBERTURAAP) & "', '"
            ssql = ssql & Trim(vgTelefono) & "', '"
            ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", '"
            ssql = ssql & Trim(vLote) & "', '"
            ssql = ssql & "0', '"
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

