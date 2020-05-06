Attribute VB_Name = "TarjetaPlata"
Option Explicit

Public Sub ImportadorTarjetaPlata()
Dim ssql As String, rsc As New Recordset, rsc2 As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset, rsrep As New Recordset, rsprod As New Recordset
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim vlineasTotales As Long
Dim sArchivo As String
Dim rsCMP As New Recordset
Dim vlog As String
Dim vArchivo As String
Dim vTInicial As Date
Dim regMod As Long
'Dim vCantDeErrores As Integer
Dim vIDCONTACTO

On Error Resume Next
    
    'SKINNER
    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos668"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatos668", cn, adOpenKeyset, adLockOptimistic
    'Marge
    cn2.Execute "DELETE FROM TM_ImportaContactosGeneralTipoCallYAsistencias"
    rsc2.Open "SELECT * FROM TM_ImportaContactosGeneralTipoCallYAsistencias", cn2, adOpenKeyset, adLockOptimistic
    
vgidCia = lIdCia '10001610
vgidCampana = lIdCampana '754
vgidUtEmpresa = 486 '  lIdUtempresaCall
vgidCampanaCall = 453 ' lIdCampanaCall

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
Set oSheet = oBook.Worksheets(1)
    
Dim filas As Long
Dim columnas As Long
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)
'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(21)
camposParaValidar(0) = "ID Cliente"
camposParaValidar(1) = "Apellido y Nombre"
camposParaValidar(2) = "ID Tipo Documento"
camposParaValidar(3) = "# Documento"
camposParaValidar(4) = "Fecha de Nacimiento"
camposParaValidar(5) = "Sexo"
camposParaValidar(6) = "Email"
camposParaValidar(7) = "Email2"
camposParaValidar(8) = "Calle"
camposParaValidar(9) = "Altura"
camposParaValidar(10) = "Piso"
camposParaValidar(11) = "Dpto"
camposParaValidar(12) = "Direccion"
camposParaValidar(13) = "Localidad"
camposParaValidar(14) = "Provincia"
camposParaValidar(15) = "CP"
camposParaValidar(16) = "Pais"
camposParaValidar(17) = "Telefono1"
camposParaValidar(18) = "Telefono2"
camposParaValidar(19) = "Telefono3"
camposParaValidar(20) = "Vigencia"
camposParaValidar(21) = "Producto Adquirido"


If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    Dim sFile As New FileSystemObject
    Dim fln As TextStream
    Set fln = sFile.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
    fln.WriteLine "Errores"
    
'======='control de lectura del archivo de datos=======================
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Exit Sub
    End If
      
'=====inicio del control de corrida====================================
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

'Dim sFileCSV As New FileSystemObject
'    oBook.Activate
'    vArchivo = Mid(FileImportacion, 1, InStr(1, FileImportacion, ".")) & "csv"
'   ' sFileCSV.OpenTextFile App.Path & vgPosicionRelativa & sDirImportacion & "\" & vArchivo
'    If sFileCSV.FileExists(App.Path & vgPosicionRelativa & sDirImportacion & "\" & vArchivo) Then
'        sFileCSV.DeleteFile (App.Path & vgPosicionRelativa & sDirImportacion & "\" & vArchivo)
'                        If Err Then
'                            MsgBox Err.Description
'                            Err.Clear
'                            Exit Sub
'                        End If
'    End If
'    ActiveWorkbook.SaveAs FileName:=
'        App.Path & vgPosicionRelativa & sDirImportacion & "\" & vArchivo, FileFormat
'        :=xlCSV, CreateBackup:=False

'=======seteo control de lote===========================================================
    Dim lLote As Long
    Dim vLote As Long
    Dim nroLinea As Long
    Dim LongDeLote As Long
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
'=======================================================================================

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
        rsc.AddNew
        rsc2.AddNew
        
'=======Control de Lote===============================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'=====================================================
        vCantDeErrores = 0
        
        For lCol = 1 To columnas
            sName = col.Item(lCol)
            v = oSheet.Cells(lRow, lCol)
            If IsEmpty(v) = False Then
            
            If lCol = 1 And IsEmpty(v) Then Exit Do
            
            vlog = ""
            
            Select Case UCase(Trim(sName))
                Case "APELLIDO Y NOMBRE"
                    rsc("APELLIDOYNOMBRE").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("apellidoynombre").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                    vgAPELLIDOYNOMBRE = v
                    

                Case "ID TIPO DOCUMENTO"
                
                
                Case "# DOCUMENTO"
                    rsc("NumeroDeDocumento") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("NroDocumento") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                    vgNumeroDeDocumento = v
                Case "FECHA DE NACIMIENTO"
                    If Len(v) > 0 And Not IsDate(v) Then
                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
                    Else
                        If IsDate(v) Then rsc("FechaNac").Value = v
                        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                          rsc2("FechaNac").Value = v
                        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                        vgFechaDeNacimiento = v
                    End If
                    
                Case "SEXO"
                    If UCase(v) = "FEMENINO" Then v = "F"
                    If UCase(v) = "MASCULINO" Then v = "M"
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Sexo").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                    vgSexo = v
                    
               Case "EMAIL"
                    rsc("email").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("email").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                    vgEmail = v
                    
                Case "EMAIL2"
                    rsc("email2").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("email2").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                    vgEmail2 = v
                    
                Case "CALLE"
                    rsc("DOMICILIO") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Calle") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgDOMICILIO = rsc("DOMICILIO")
                    
                    
                Case "ALTURA"
                    rsc("DOMICILIO") = rsc("DOMICILIO") & " " & v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Altura") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgDOMICILIO = rsc("DOMICILIO")
                Case "PISO"
                    rsc("DOMICILIO") = rsc("DOMICILIO") & " Piso " & v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Piso") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgDOMICILIO = rsc("DOMICILIO")
                Case "DPTO"
                    rsc("DOMICILIO") = rsc("DOMICILIO") & " dpto " & v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Dpto") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgDOMICILIO = rsc("DOMICILIO")
                    
                Case "LOCALIDAD"
                    rsc("LOCALIDAD") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("LOCALIDAD") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgLOCALIDAD = v
                    
                Case "PROVINCIA"
                    rsc("PROVINCIA") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("PROVINCIA") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgPROVINCIA = v
                    
                Case "CP"
                    rsc("CODIGOPOSTAL") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("CP") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgCODIGOPOSTAL = v
                    
                Case "PAIS"
                    rsc("Pais") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Pais") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgPais = v
                    
                Case "TELEFONO1"
                    rsc("Telefono") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Telefono") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgTelefono = v
                    
                Case "TELEFONO2"
                    If InStr(1, v, "E+") > 0 Then
                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
                    Else
                        rsc("Telefono2") = v
                        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                       rsc2("Telefono2") = v
                       vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                       vgTelefono2 = v
                       
                    End If
                    
                Case "TELEFONO3"
                    rsc("Telefono3") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("Fax") = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgTelefono3 = v
                    
'                Case "TIPO DE PRESTAMO"
'                    rsc("COBERTURAVEHICULO") = v
'                    rsc("COBERTURAVIAJERO") = v
'                    rsc("COBERTURAHOGAR") = v
'                    rsc2("CUSTOM1").Value = v
'
'                Case "PRODUCTO"
'                    rsc("COBERTURAVEHICULO") = v
'                    rsc("COBERTURAVIAJERO") = v
'                    rsc("COBERTURAHOGAR") = v
'                    rsc2("CUSTOM6").Value = v
'
'                Case "NUMERO DE CREDITO"
'                    rsc("NroPoliza").Value = v
'                    rsc2("CUSTOM2").Value = v
'                    ssql = "Select idcontacto From tm_contactos where idcampana = " & rsc2("IdCampana") & " and CUSTOM2 = '" & v & "'"
'                    'rsrep.Open ssql, cn2, adOpenForwardOnly, adLockReadOnly
                    'If Not rsrep.EOF Then
                    '    vlog = vlog & " Prestamo existente en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
                    'End If
                    'rsrep.Close
                    
                Case "ID CLIENTE"
                    rsc("NroPoliza").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    rsc2("CUSTOM2").Value = v
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                    vgNROPOLIZA = v
                   ' ssql = "Select idcontacto From tm_contactos where idcampana = " & rsc2("IdCampana") & " and CUSTOM2 = '" & v & "'"
                    'rsrep.Open ssql, cn2, adOpenForwardOnly, adLockReadOnly
                    'If Not rsrep.EOF Then
                    '    vlog = vlog & " Prestamo existente en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
                    'End If
                    'rsrep.Close
   
               ' Case "CAPITAL PEDIDO"
                '    rsc2("CUSTOM3").Value = v
                'Case "COMISION TEORICA"
               '     rsc2("CUSTOM4").Value = v
               
'                Case "FECHA DE SOLICITUD"
'                    If Len(v) > 0 And Not IsDate(v) Then
'                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
'                    Else
'                        rsc("FECHAVENCIMIENTO").Value = v
'                        rsc2("CUSTOM5").Value = v
'                    End If
'
'                Case "VIGENCIA TOPE DE PRESTACION"
'                    If Len(v) > 0 And Not IsDate(v) Then
'                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
'                    Else
'                        rsc("FECHAVENCIMIENTO").Value = v
'                        rsc2("CUSTOM5").Value = v
'                    End If
                    
                Case "VIGENCIA"
                    If Len(v) > 0 And Not IsDate(v) Then
                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
                    Else
                        rsc("FECHAVENCIMIENTO").Value = v
                        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                        rsc2("CUSTOM5").Value = v
                        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                        vgFECHAVENCIMIENTO = v
                        
                    End If
                    
'                Case "ALTA DE PRESTACION"
'                    If Len(v) > 0 And Not IsDate(v) Then
'                        vlog = vlog & " Error en linea " & lRow & " en el campo " & sName & Chr(10) & Chr(13)
'                    Else
'                        rsc("FECHAVIGENCIA").Value = v
'                        rsc2("CUSTOM7").Value = v
'                    End If
                    
                Case "PRODUCTO ADQUIRIDO"

                    If Len(v) <= 2 Then
                     ssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & vgidCampana & " and idproductoencliente ='" & v & "'"
                     rsprod.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                        If Not rsprod.EOF Then
                             rsc("COBERTURAVEHICULO") = rsprod("coberturavehiculo")
                             vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                             
                             rsc("COBERTURAVIAJERO") = rsprod("coberturaviajero")
                             vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                             
                             rsc("COBERTURAHOGAR") = rsprod("coberturahogar")
                             vgCOBERTURAHOGAR = rsprod("coberturahogar")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                             
                             rsc2("custom6") = rsprod("descripcion")
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampanaCall, "", lRow, sName)
                             rsc("IdProducto") = v
                             vgIdProducto = v
                             vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
                        End If
                     rsprod.Close
                    End If
                    
                'Case "OTORGA"
                 '   rsc("Conductor") = v
                
            End Select
            
            End If
        Next
          
        vgDOMICILIO = rsc("DOMICILIO")
        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
        'vgCORRIDA = rsc("CORRIDA")
        rsc("CORRIDA") = vgCORRIDA
        If Not IsDate(rsc("FECHAVIGENCIA").Value) Then rsc("FECHAVIGENCIA").Value = Now
        If Not IsDate(rsc2("CUSTOM7").Value) Then rsc2("CUSTOM7").Value = Now
        vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "", lRow, sName)
        
        
        
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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'" ' and Nrosecuencial = '" & vgNROSECUENCIAL & "'"
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    Dim vdif As Long
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                'If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                'If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
                If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then vdif = vdif + 1
                If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then vdif = vdif + 1
                If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then vdif = vdif + 1
                If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then vdif = vdif + 1
                If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then vdif = vdif + 1
                If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then vdif = vdif + 1
                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                If Trim(rscn1("Gama")) <> Trim(vgGama) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
                
                If vdif > 0 Then
                vdif = vdif
                End If
                
            End If
        rscn1.Close
'=============================================================================================================
    Dim rscn2 As New Recordset
    ssql = "select * from UniversalT.dbo.tm_Contactos  where IDUTEMPRESA = " & vgidUtEmpresa & " and CUSTOM2 = '" & rsc2("CUSTOM2").Value & "'"
    rscn2.Open ssql, cn2, adOpenForwardOnly, adLockReadOnly
     
    If Not rscn2.EOF Then
        rsc2("IDCONTACTO") = rscn2("IDCONTACTO")
        rsc2("Modificaciones") = vdif
        rsc2("idLote") = vLote
    Else
        rsc2("IDCONTACTO") = 0
        rsc2("Modificaciones") = 1
        rsc2("idLote") = vLote
    End If
    
    rscn2.Close
'==============================================================================================================
    rsc("IDPOLIZA") = vgIDPOLIZA
    rsc("Modificaciones") = vdif
    rsc("idLote") = vLote
     
    rsc.Update
    rsc2.Update
'========Control de errores==============================================================
        If Err Then
            vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow - 1, "")
            Err.Clear
        
        End If
'========================================================================================
        If vdif > 0 Then
            regMod = regMod + 1
        End If
        
        lRow = lRow + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow
                ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lIdCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
            ll100 = 0
        End If
        DoEvents
    Loop

'    oExcel.Workbooks.Close
'    Set oExcel = Nothing
'================Control de Leidos=======================================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
    listoParaProcesar
    
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos===========================================
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'========================================================================================
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea 1" & Chr(13) & " de " & lRow & " Procesando los datos"
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 600
        cn1.Execute "TM_CargaPolizasTarjetaPlataControlado " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
        cn2.Execute "Importacion_Contactos_General_TipoCallyAsistencias" & " " & lLote & ", " & vgidUtEmpresa & ", " & vgidCampanaCall
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If rsCMP("UltimaCorridaError") <> "OK" Then
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            'lLote = vLote + 1 'para salir del FOR
        Else
            ImportadordePolizas.txtprocesando.Text = "Procesando el Lote " & lLote & " De " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " Procesando los datos" & Chr(13) & " Datos Procesados Hasta el Momento " & rsCMP("UltimaCorridaCantidadImportada")
            DoEvents
        End If
        rsCMP.Close

'vTInicial = Now()
'Do Until DateDiff("s", vTInicial, Now()) > 60
'
'    vLote = vLote
'Loop
    Next lLote
    

'=====================================================================================
'    ImportadordePolizas.txtProcesando.Text = "Importando copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
'    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    cn1.Execute "TM_CargaPolizasTipoCallYAsistenciasExcel 754, 10001610"
'    cn2.Execute "Importacion_Contactos_General_TipoCallyAsistencias" & " " & lLote
    
    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & vgidCampana & ", " & vgCORRIDA
    Procesado
'=====================================================================================
    ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & lRow & " FinDeProceso"
    ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF

    
    
Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

oExcel.Workbooks.Close
Set oExcel = Nothing


'Exit Sub

'errores:
'
'        oExcel.Workbooks.Close
'    Set oExcel = Nothing
'    vgErrores = 1
'    MsgBox "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
'    If lRow = 0 Then
'        MsgBox Err.Description
'    Else
'        MsgBox Err.Description & " en linea " & lRow
'    End If
'
End Sub


