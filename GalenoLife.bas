Attribute VB_Name = "GalenoLife"
Option Explicit

Public Sub ImportarGalenoLife()

Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim ssqlProv As String
Dim rsCMPProv As New Recordset
Dim rsp As New Recordset
Dim rsprod As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim vlineasTotales As Long
Dim sArchivo As String
Dim vHoja As String
Dim lRowH As Long
Dim vtotalDeLineas As Long
Dim vdir As Integer
Dim regMod As Long
Dim vVigenciaVigente As String

'Variables para Colon Seguros
Dim vDepto As String
Dim vCalle As String
Dim vAltura As String
Dim vApellido As String
Dim vNombre As String
Dim vNro As String
Dim vPiso As String
Dim Ll As Long
Dim vFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim vLinea As Long
Dim vPosicion As Long
Dim vCampo As String
'Dim vCoveruraVidrios As Double
Dim fechaActual As Date

fechaActual = Now

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

TablaTemporal

On Error Resume Next

'================
Dim vCS As String
vCS = ";"
'================

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

Ll = 2
ll100 = 0
'vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & "Chubb.txt"
vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
If Not fs.FileExists(vFile) Then Exit Sub
Set tf = fs.OpenTextFile(vFile, ForReading, True)
'======='control de lectura del archivo de datos=======================
If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If
'=====inicio del control de corrida====================================
Dim rsCorr As New Recordset
cn1.Execute "TM_CargaPolizasLogDeSetCorridasxcia " & lIdCia & ", " & vgCORRIDA
ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcia = " & lIdCia & " and Registrosleidos is null"
rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
If rsCorr.EOF Then
    MsgBox "no se determino la corrida, se detendra el proceso"
    Exit Sub
Else
    vgCORRIDA = rsCorr("corrida")
End If
'=======seteo control de lote================================================================

Dim lLote As Long
Dim vLote As Long
Dim nroLinea As Long
Dim LongDeLote As Long
Dim vControlDeModificados As Long
LongDeLote = 1000
nroLinea = 1
vLote = 1
'=================================================================================================

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If
    tf.SkipLine 'Saltea encabezado
   Do Until tf.AtEndOfStream
        vLinea = Ll
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "")
          
            '====maneja los lotes para corte de importacion========
            nroLinea = nroLinea + 1
            
  '          If nroLinea = 13 Then
   '             MsgBox 1
    '        End If
            
            If nroLinea = LongDeLote + 1 Then
                vLote = vLote + 1
                vControlDeModificados = 0
                nroLinea = 1
            End If
            '======================================================
        Blanquear
        vPosicion = 0
        vVigenciaVigente = "00:00:00"

    '==================================================================================================
        vCampo = "Asegurado Nombre"
        vPosicion = vPosicion + 1
            vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Asegurado Domicilio"
        vPosicion = vPosicion + 1
            vgDOMICILIO = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Localidad"
        vPosicion = vPosicion + 1
            vgLOCALIDAD = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Provincia"
        vPosicion = vPosicion + 1
            vgPROVINCIA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Asegurado Tipo Doc."
        vPosicion = vPosicion + 1
            vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Asegurado Nro. Doc."
        vPosicion = vPosicion + 1
            vgNumeroDeDocumento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Asegurado Fec.Nacim."
        vPosicion = vPosicion + 1
        vgFechaDeNacimiento = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        If Not IsDate(vgFechaDeNacimiento) Then
            vgFechaDeNacimiento = Now
            vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
        End If
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Póliza"
        vPosicion = vPosicion + 1
            vgNROPOLIZA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Tipo Pza."
        vPosicion = vPosicion + 1
            vgCodigoDeProductor = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "IDPRODUCTO"
        vPosicion = vPosicion + 1
            v = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            If Len(v) > 0 Then
            ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente = '" & v & "'"
            rsprod.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                If Not rsprod.EOF Then
                     vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCOBERTURAHOGAR = rsprod("coberturahogar")
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                     vgCodigoEnCliente = v
                     vgTipodeOperacion = v
                     vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", Ll, vCampo)
                Else
                     vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", Ll, vCampo)
                
                End If
            rsprod.Close
            End If
        'vgCodigoEnCliente = sLine 'Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Adjunto"
        vPosicion = vPosicion + 1
            vgAgencia = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "Mail"
        vPosicion = vPosicion + 1
            vgEmail = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    '==================================================================================================
        vCampo = "MontoCoverturaVidrios"
        vPosicion = vPosicion + 1
            vgTopeCritales = sLine 'Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        'sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)

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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and NumeroDeDocumento = '" & Trim(vgNumeroDeDocumento) & "' and TipodeOperacion = '" & Trim(vgTipodeOperacion) & "' "
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    
    vgFECHAVENCIMIENTO = DateAdd("m", 3, fechaActual)
    vVigenciaVigente = (rscn1("FECHAVIGENCIA"))
    If Err.Number = 3021 Then 'Limpio error: El valor de BOF o EOF es True, o el actual registro se elimino
        Err.Clear
    End If
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vVigenciaVigente) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
                If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                If Trim(rscn1("CodigoEnCliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
                If Trim(rscn1("MontoCoverturaVidrios")) <> Trim(vgTopeCritales) Then vdif = vdif + 1
                If Trim(rscn1("Agencia")) <> Trim(vgAgencia) Then vdif = vdif + 1
                If Trim(rscn1("Email")) <> Trim(vgEmail) Then vdif = vdif + 1
                If vgCOBERTURAHOGAR <> "" Then
                    If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                End If
                If vgCOBERTURAVEHICULO <> "" Then
                    If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                End If
                If vgCOBERTURAVIAJERO <> "" Then
                    If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                End If
                vgIDPOLIZA = rscn1("idpoliza")
            End If
            
        If vgIDPOLIZA = 0 Then
            vVigenciaVigente = fechaActual
        End If
        
        rscn1.Close
'-=================================================================================================================
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
            ssql = ssql & "IdPoliza, "
            ssql = ssql & "CodigoEnCliente, "
            ssql = ssql & "IdCampana, "
            ssql = ssql & "idcia, "
            ssql = ssql & "NROPOLIZA, "
            ssql = ssql & "APELLIDOYNOMBRE, "
            ssql = ssql & "NumeroDeDocumento, "
            ssql = ssql & "Email, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "TipodeOperacion, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
            ssql = ssql & "FECHAVIGENCIA, "
            ssql = ssql & "FECHAVENCIMIENTO, "
            ssql = ssql & "FechadeNacimiento, "
            ssql = ssql & "CORRIDA, "
            ssql = ssql & "IdLote, "
            ssql = ssql & "MontoCoverturaVidrios, "
            ssql = ssql & "Modificaciones)"
            
            ssql = ssql & " values("
            ssql = ssql & Trim(vgIDPOLIZA) & ", '"
            ssql = ssql & Trim(vgCodigoEnCliente) & "', "
            ssql = ssql & Trim(vgidCampana) & ", "
            ssql = ssql & Trim(vgidCia) & ", '"
            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
            ssql = ssql & Trim(vgEmail) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
            ssql = ssql & Trim(vVigenciaVigente) & "', '"
            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
            ssql = ssql & Trim(vgFechaDeNacimiento) & "', "
            ssql = ssql & Trim(vgCORRIDA) & ", '"
            ssql = ssql & Trim(vLote) & "', "
            ssql = ssql & Trim(vgTopeCritales) & ", '"
            ssql = ssql & Trim(vdif) & "') "
            cn.Execute ssql
            
        
'========Control de errores=========================================================
        If Err Then
            vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, vCampo)
            'MsgBox "PARA"
            Err.Clear
        
        End If
'====================================================================================

         If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
        ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================

                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lidCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
            ll100 = 0
        End If
        DoEvents
    Loop

'================Control de Leidos================================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
    listoParaProcesar
'=================================================================================
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos====================================
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'=================================================================================
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

    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & vgidCampana & ", " & vgCORRIDA
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


Public Sub ImportarGalenoLifeOld()

Dim ssql As String, rsc As New Recordset
Dim sssql As String
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vIDCampana As Long
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim rsprod As New Recordset
Dim vlineasTotales As Long
Dim sArchivo As String
Dim regMod As Long
Dim vSucursal As String


On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

TablaTemporal
 
Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application
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

Dim camposParaValidar(3)
camposParaValidar(0) = "DOCUMENTO"
camposParaValidar(1) = "IDPRODUCTO"
camposParaValidar(2) = "APELLIDOYNOMBRE"


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

'====='Control de Lote===================================================
        nroLinea = nroLinea + 1
        If lRow = 6195 Then
            MsgBox "BOSXX"
        End If
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
                Case "DOMICILIO"
                    vgDOMICILIO = Replace(v, "'", "")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "LOCALIDAD"
                    vgLOCALIDAD = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PROVINCIA"
                    vgPROVINCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPODOCUMENTO"
                    vgTipodeDocumento = Mid(v, 1, 20)
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DOCUMENTO"
                    vgNumeroDeDocumento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHANACIMIENTO"
                    If Not IsDate(vgFechaDeNacimiento) Then
                        vgFechaDeNacimiento = "00:00:00"
                    End If
                    vgFechaDeNacimiento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "NROPOLIZA"
                    vgNROPOLIZA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TIPOPOLIZA"
                    vgTipodeOperacion = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "IDPRODUCTO"
                    If Len(v) > 0 Then
                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente =  " & v
                     rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly
                        If Not rsprod.EOF Then
                             vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgCOBERTURAHOGAR = rsprod("coberturahogar")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                             vgCodigoEnCliente = v
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                        Else
                             vCantDeErrores = vCantDeErrores + LoguearErrorDeConcepto("Producto Inexistente", flnErr, vgidCampana, "", lRow, sName)
                        
                        End If
                     rsprod.Close
                    End If

                Case "MAIL"
                   vgEmail = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
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
         ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and NumeroDeDocumento = '" & Trim(vgNumeroDeDocumento) & "'"
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
            vgIDPOLIZA = 0
                    If Not rscn1.EOF Then
                        vdif = 0  'setea la variale de control de repetido con modificacion en cero
                        If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                        If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                        If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                        If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                        If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                        If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                        If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                        If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                        If Trim(rscn1("CodigoEnCliente")) <> Trim(vgCodigoEnCliente) Then vdif = vdif + 1
                        If Trim(rscn1("Email")) <> Trim(vgEmail) Then vdif = vdif + 1
                        If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
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
            ssql = ssql & "Email, "
            ssql = ssql & "DOMICILIO, "
            ssql = ssql & "TipodeDocumento, "
            ssql = ssql & "TipodeOperacion, "
            ssql = ssql & "COBERTURAVEHICULO, "
            ssql = ssql & "COBERTURAVIAJERO, "
            ssql = ssql & "COBERTURAHOGAR, "
            ssql = ssql & "LOCALIDAD, "
            ssql = ssql & "PROVINCIA, "
            ssql = ssql & "FechadeNacimiento, "
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
            ssql = ssql & Trim(vgEmail) & "', '"
            ssql = ssql & Trim(vgDOMICILIO) & "', '"
            ssql = ssql & Trim(vgTipodeDocumento) & "', '"
            ssql = ssql & Trim(vgTipodeOperacion) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
            ssql = ssql & Trim(vgPROVINCIA) & "', '"
            ssql = ssql & Trim(vgFechaDeNacimiento) & "', "
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


Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

oExcel.Workbooks.Close
Set oExcel = Nothing

End Sub

