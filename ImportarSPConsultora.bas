Attribute VB_Name = "SPConsultora"
Option Explicit

Public Sub ImportarSPConsultora()

Dim sssql As String, rsc As New Recordset
Dim lCol As Long
Dim lRow As Long
Dim ll100 As Long
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
Dim NombreDividido() As String
Dim PrimeraParte As String

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

TablaTemporal

'================
'lRow = 2
'lCol = 1
'll100 = 0
''================
''=======control de lote============
'LongDeLote = 1000
'nroLinea = 1
'vLote = 1
'================
'
'Dim rsCorr As New Recordset
'cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lIdCampana & ", " & vgCORRIDA
'ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcampana = " & lIdCampana & " and Registrosleidos is null"
'rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'If rsCorr.EOF Then
'    MsgBox "no se determino la corrida, se detendra el proceso"
'    Exit Sub
'Else
'    vgCORRIDA = rsCorr("corrida")
'End If

'NombreDividido() = Split(fileimportacion, ".")
'PrimeraParte = NombreDividido(1)
'
'LeerArchivo fileimportacion, vgCORRIDA, lRow, ll100, vLote, LongDeLote, lCol, vlineasTotales
''================
''lRow = 2
''lCol = 1
''================
'LeerArchivo "SP_ComodoroRivadavia.xlsx", vgCORRIDA, lRow, ll100, vLote, LongDeLote, lCol, vlineasTotales


Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application
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

Dim camposParaValidar(7)
camposParaValidar(0) = "APELLIDOYNOMBRE"
camposParaValidar(1) = "DOCUMENTO"
camposParaValidar(2) = "INICIOVIGENCIA"
camposParaValidar(3) = "FINVIGENCIA"
camposParaValidar(4) = "IDPRODUCTO"
camposParaValidar(5) = "PROVINCIA"
camposParaValidar(6) = "LOCALIDAD"

'================================objeto excel para almacenar errores=======================================
If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    Dim vCantDeErrores As Integer
    Dim sFileErr As New FileSystemObject
    Dim flnErr As TextStream
    Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
    flnErr.WriteLine "Errores"
    vCantDeErrores = 0
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
'=============================== incializacion de variables de control de lote========================================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    ll100 = 0
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
    '=======Control de Lote===============================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
    '===============Comienzo de lectura del excel======================================
        Blanquear
        vCantDeErrores = 0

        For lCol = 1 To columnas
            sName = col.Item(lCol)
            v = oSheet.Cells(lRow, lCol)
            If IsEmpty(v) = False Then

            If lCol = 1 And IsEmpty(v) Then Exit Do
            'vlog = ""'indica en que linea se produce el error en la lectura
            'sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "APELLIDOYNOMBRE"
                    vgAPELLIDOYNOMBRE = Replace(v, "'", "")
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DOCUMENTO"
                    vgNumeroDeDocumento = v
                    vgNROPOLIZA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "TELEFONO"
                    vgTelefono = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "INICIOVIGENCIA"
                    vgFECHAVIGENCIA = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FINVIGENCIA"
                   vgFECHAVENCIMIENTO = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "FECHANACIMIENTO"
                    vgFechaDeNacimiento = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "DOMICILIO"
                    vgDOMICILIO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "CODIGOPOSTAL"
                    vgCODIGOPOSTAL = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "VEHICULO"
                    vgMARCADEVEHICULO = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "IDPRODUCTO"
                    If Len(v) > 0 Then
                     sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lidCampana & "  and idproductoencliente = " & v
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
                        End If
                     rsprod.Close
                    End If
                Case "PATENTE"
                    vgPATENTE = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "PROVINCIA"
                   vgPROVINCIA = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
                Case "LOCALIDAD"
                   vgLOCALIDAD = v
                   vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
            End Select
            End If
        Next
        
    '=============  IMPORTANTE   ================================================================
    Dim vcamp As Integer
    Dim vdif As Long
    Dim rscn1 As New Recordset
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' " & " and patente = '" & Trim(vgPATENTE) & "' "
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If vgFECHAVIGENCIA Then
                    If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                End If
                If vgFECHAVENCIMIENTO Then
                    If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                End If
                If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If vgCOBERTURAVEHICULO <> "" Then
                    If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                End If
                If vgCOBERTURAVIAJERO <> "" Then
                    If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                End If
                If vgCOBERTURAHOGAR <> "" Then
                    If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                End If
                If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")

            End If
        rscn1.Close
        
    '=================================================================================================================
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
        ssql = ssql & "IdPoliza, "
        ssql = ssql & "CodigoEnCliente, "
        ssql = ssql & "IdCampana, "
        ssql = ssql & "idcia, "
        ssql = ssql & "NROPOLIZA, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "Telefono, "
        ssql = ssql & "CODIGOPOSTAL, "
        ssql = ssql & "FechadeNacimiento, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "NUMERODEDOCUMENTO, "
        ssql = ssql & "FECHAVENCIMIENTO, "
        ssql = ssql & "APELLIDOYNOMBRE, "
        ssql = ssql & "COBERTURAVEHICULO, "
        ssql = ssql & "COBERTURAVIAJERO, "
        ssql = ssql & "COBERTURAHOGAR, "
        ssql = ssql & "MARCADEVEHICULO, "
        ssql = ssql & "PATENTE, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "CORRIDA, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "Modificaciones)"

        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
        ssql = ssql & Trim(vgIdProducto) & ", "
        ssql = ssql & Trim(vgidCampana) & ", "
        ssql = ssql & Trim(vgidCia) & ", '"
        ssql = ssql & Trim(vgNROPOLIZA) & "', '"
        ssql = ssql & Trim(vgLOCALIDAD) & "', '"
        ssql = ssql & Trim(vgPROVINCIA) & "', '"
        ssql = ssql & Trim(vgTelefono) & "', '"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
        ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
        ssql = ssql & Trim(vgPATENTE) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', "
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

    '================Control de Leidos===============================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
    listoParaProcesar
    '=================================================================================
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
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & lIdCia & ", " & lidCampana
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
    '=====================================================================================
    ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " FinDeProceso"
    ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF


Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

oExcel.Workbooks.Close
Set oExcel = Nothing

End Sub
    
'    Public Function LeerArchivo(fileimportacion As String, vgCORRIDA As Long, lRow As Long, ll100 As Long, vLote As Integer, LongDeLote As Long, lCol As Long, vlineasTotales As Long)
'
'    Dim sssql As String, rsc As New Recordset
'    Dim v, sName
'    Dim ssql As String
'    Dim rsprod As New Recordset
'    Dim regMod As Long
'    Dim nroLinea As Integer
'    On Error Resume Next
'
'    Dim col As New Scripting.Dictionary
'    Dim oExcel As Excel.Application
'    Dim oBook As Excel.Workbook
'    Dim oSheet As Excel.Worksheet
'    Set oExcel = New Excel.Application
'    oExcel.Visible = False
'    Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
'    Set oSheet = oBook.Worksheets(1)
'
'    'array para leer la primera row del excel y cargar los campos que trae el excel.
'    Dim filas As Integer
'    Dim columnas As Integer
'    Dim extremos(1)
'    columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
'    extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)
'    filas = extremos(1)
'
'    Dim camposParaValidar(7)
'    camposParaValidar(0) = "APELLIDOYNOMBRE"
'    camposParaValidar(1) = "DOCUMENTO"
'    camposParaValidar(2) = "INICIOVIGENCIA"
'    camposParaValidar(3) = "FINVIGENCIA"
'    camposParaValidar(4) = "IDPRODUCTO"
'    camposParaValidar(5) = "PROVINCIA"
'    camposParaValidar(6) = "LOCALIDAD"
'
'    '================================objeto excel para almacenar errores=======================================
'    If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then
'
'        Dim vCantDeErrores As Integer
'        Dim sFileErr As New FileSystemObject
'        Dim flnErr As TextStream
'        Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
'        flnErr.WriteLine "Errores"
'        vCantDeErrores = 0
'    '======='control de lectura del archivo de datos
'        If Err Then
'            MsgBox Err.Description
'            Err.Clear
'            Exit Function
'        End If
'
'    '=============================== incializacion de variables de control de lote========================================
'
''        Dim vLote As Long
''        Dim nroLinea As Long
''        Dim LongDeLote As Long
''        LongDeLote = 1000
'        nroLinea = 1
'        lRow = 2
'        lCol = 1
''        vLote = 1
''        lRow = 2
''        lCol = 1
'
'        Do While lCol < columnas + 1
'            v = oSheet.Cells(1, lCol)
'            If IsEmpty(v) Then Exit Do
'            sName = v
'            col.Add lCol, v
'            lCol = lCol + 1
'        Loop
'
'        Do While lRow < filas + 1
'    '=======Control de Lote===============================
'            nroLinea = nroLinea + 1
'            If nroLinea = LongDeLote + 1 Then
'                vLote = vLote + 1
'                nroLinea = 1
'            End If
'    '===============Comienzo de lectura del excel======================================
'                vCantDeErrores = 0
'            For lCol = 1 To columnas
'                sName = col.Item(lCol)
'                v = oSheet.Cells(lRow, lCol)
'                If IsEmpty(v) = False Then
'
'                If lCol = 1 And IsEmpty(v) Then Exit Do
'                Select Case UCase(Trim(sName))
'                    Case "APELLIDOYNOMBRE"
'                        vgAPELLIDOYNOMBRE = Replace(v, "'", "´")
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "DOCUMENTO"
'                        vgNumeroDeDocumento = v
'                        vgNROPOLIZA = v
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "INICIOVIGENCIA"
'                        vgFECHAVIGENCIA = v
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "FINVIGENCIA"
'                       vgFECHAVENCIMIENTO = v
'                       vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "DOMICILIO"
'                        vgDOMICILIO = v
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "VEHICULO"
'                        vgMARCADEVEHICULO = v
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "IDPRODUCTO"
'                        If Len(v) > 0 Then
'                         sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & lIdCampana & "  and idproductoencliente = " & v
'                         rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly
'                            If Not rsprod.EOF Then
'                                 vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
'                                 vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                                 vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
'                                 vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                                 vgCOBERTURAHOGAR = rsprod("coberturahogar")
'                                 vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                                 vgIdProducto = v
'                                 vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                            End If
'                         rsprod.Close
'                        End If
'                    Case "PATENTE"
'                        vgPATENTE = v
'                        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "PROVINCIA"
'                       vgPROVINCIA = v
'                       vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                    Case "LOCALIDAD"
'                       vgLOCALIDAD = v
'                       vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
'                End Select
'                End If
'            Next
'
'     Dim vcamp As Integer
'     Dim vdif As Long
'     Dim rscn1 As New Recordset
'        ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lIdCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
'        rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'        vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
'        vgIDPOLIZA = 0
'                If Not rscn1.EOF Then
'                    vdif = 0  'setea la variale de control de repetido con modificacion en cero
'                    If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
'                    If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
'                    If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
'                    If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
'                    If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
'                    If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
'                    If vgFECHAVIGENCIA Then
'                        If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
'                    End If
'                    If vgFECHAVENCIMIENTO Then
'                        If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
'                    End If
'                    If Trim(rscn1("DOCUMENTO")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
'                    If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
'                    If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
'                    If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
'                    If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
'                    If Trim(rscn1("CodigoEnCliente")) <> Trim(vgIdProducto) Then vdif = vdif + 1
'                    vgIDPOLIZA = rscn1("idpoliza")
'                End If
'            rscn1.Close
'    '-=================================================================================================================
'            ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & vgidCampana & "("
'            ssql = ssql & "IdPoliza, "
'            ssql = ssql & "CodigoEnCliente, "
'            ssql = ssql & "IdCampana, "
'            ssql = ssql & "idcia, "
'            ssql = ssql & "NROPOLIZA, "
'            ssql = ssql & "LOCALIDAD, "
'            ssql = ssql & "PROVINCIA, "
'            ssql = ssql & "FECHAVIGENCIA, "
'            ssql = ssql & "NUMERODEDOCUMENTO, "
'            ssql = ssql & "FECHAVENCIMIENTO, "
'            ssql = ssql & "APELLIDOYNOMBRE, "
'            ssql = ssql & "COBERTURAVEHICULO, "
'            ssql = ssql & "COBERTURAVIAJERO, "
'            ssql = ssql & "COBERTURAHOGAR, "
'            ssql = ssql & "MARCADEVEHICULO, "
'            ssql = ssql & "PATENTE, "
'            ssql = ssql & "DOMICILIO, "
'            ssql = ssql & "CORRIDA, "
'            ssql = ssql & "IdLote, "
'            ssql = ssql & "Modificaciones)"
'            ssql = ssql & " values("
'            ssql = ssql & Trim(vgIDPOLIZA) & ", "
'            ssql = ssql & Trim(vgIdProducto) & ", "
'            ssql = ssql & Trim(vgidCampana) & ", "
'            ssql = ssql & Trim(vgidCia) & ", '"
'            ssql = ssql & Trim(vgNROPOLIZA) & "', '"
'            ssql = ssql & Trim(vgLOCALIDAD) & "', '"
'            ssql = ssql & Trim(vgPROVINCIA) & "', '"
'            ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
'            ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
'            ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
'            ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
'            ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
'            ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
'            ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
'            ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
'            ssql = ssql & Trim(vgPATENTE) & "', '"
'            ssql = ssql & Trim(vgDOMICILIO) & "', "
'            ssql = ssql & Trim(vgCORRIDA) & ", '"
'            ssql = ssql & Trim(vLote) & "', '"
'            ssql = ssql & Trim(vdif) & "') "
'            cn.Execute ssql
'
'
'    '========Control de errores=========================================================
'            If Err Then
'                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
'                Err.Clear
'
'            End If
'    '===========================================================================================
'
'             If vdif > 0 Then
'                regMod = regMod + 1
'            End If
'
'            lRow = lRow + 1
'            ll100 = ll100 + 1
'            If ll100 = 100 Then
'                ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow
'
'    ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
'                    ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lIdCampana & "and corrida =" & vgCORRIDA
'                    cn1.Execute ssql
'                ll100 = 0
'            End If
'            DoEvents
'        Loop
'
'
'    vlineasTotales = lRow + vlineasTotales
'
'    oExcel.Workbooks.Close
'    Set oExcel = Nothing
'
'    Else
'        MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
'    End If
'
'
'
'    End Function


'=======================================================================================================================================================================================================================================================================================================
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'=======================================================================================================================================================================================================================================================================================================
