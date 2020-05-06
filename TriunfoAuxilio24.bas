Attribute VB_Name = "TriunfoAuxilio24_OLVIDADO"
Option Explicit
Public Sub ImportarAuxilio24()
Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vCtrolIdentificacion As Boolean, vCtrolVigencia As Boolean, vCtrolVencimiento As Boolean, vCtroltipodeasistencia As Boolean, vCtrolidproductoasistencia As Boolean
Dim col As New Scripting.Dictionary
Dim vLinea As String
Dim vCertificado As String
Dim vDocumento As String
Dim vPatente As String
Dim vError As Integer
Dim vhorainicio As Date
Dim vNroMotor As String
Dim vProductosPosibles As String

Dim rscn1 As New Recordset
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim Ll As Long
Dim nroLinea As Long
Dim vCampo As String
Dim vPosicion As Long
Dim lLote As Long
Dim vLote As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
Dim rsCMP As New Recordset
Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim vLTIPODEVEHICULO As String
Dim vTipodeServicio As String
Dim vTipodeServicioActual As String

Dim vOBSERVACIONES As String
vhorainicio = Now
'Dim mExcel As New Excel.Application
'Dim wb
'        Dim oExcel As Excel.Application
'        Dim oBook As Excel.Workbook
'        Dim oSheet As Excel.Worksheet

'On Error GoTo errores
cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatosAuxilio24"
On Error Resume Next
vgidCia = lIdCia
vgidCampana = lIdCampana

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If



Dim sFile As New FileSystemObject
Dim fln As TextStream
Set fln = sFile.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion & Year(Now) & Month(Now) & Day(Now) & ".log", True)
fln.WriteLine "Linea; Certificado; Documento; Patente; Fecha; Detalle"
        
        
        
        
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
                
    v = " "
    vCtrolIdentificacion = False
    vCtrolVencimiento = False
    vCtrolVigencia = False
    vCtroltipodeasistencia = False
    vCtrolidproductoasistencia = False
    
    vgidCia = 10000546
    
    lCol = 1
    lRow = 1
    Do While lCol < 80
        v = Trim(UCase(oSheet.Range(mToChar(lCol - 1) & "1").Value))
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
        Select Case UCase(v)
            Case "PATENTE"
                vCtrolIdentificacion = True
            Case "VIGENCIAFIN"
                vCtrolVencimiento = True
            Case "FECHAEMISION"
                vCtrolVigencia = True
            Case "TIPODEASISTENCIA"
                vCtroltipodeasistencia = True
            Case "IDPRODUCTOASISTENCIA"
                vCtrolidproductoasistencia = True
            
        End Select
    Loop
    lCantCol = lCol

    If vCtrolIdentificacion = False Or vCtrolVencimiento = False Or vCtrolVigencia = False Or vCtroltipodeasistencia = False Or vCtrolidproductoasistencia = False Then
        MsgBox "Falta alguna Columna Obligatoria o esta mal la descripcion"
        Exit Sub
    End If

    If lCol = 1 Then
        MsgBox "Faltan campos"
        Exit Sub
    End If
    
    vProductosPosibles = ""
    ssql = "Select Distinct categoria from tm_Servicioxcia where idcia = " & vgidCia
    rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    vProductosPosibles = rs("Categoria")
    rs.MoveNext
    Do Until rs.EOF
        vProductosPosibles = vProductosPosibles & "," & rs("Categoria")
        rs.MoveNext
    Loop
    

'    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatosAuxilio24", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 79000
        vLinea = ""
        vCertificado = ""
        vDocumento = ""
        vPatente = ""
        vError = 0
        rsc.AddNew
        vLinea = lRow
        For lCol = 1 To lCantCol
            'v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            v = Trim(oSheet.Cells(lRow, lCol))
            If IsEmpty(v) Then v = ""
            If lCol = 1 And IsEmpty(v) Then
                rsc.CancelUpdate
                    Exit Do
                End If
            
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
            Case UCase("tipodeasistencia")
                'If v <> "MOTOS" And v <> "AUTOS" Then
                 '   rsc.CancelUpdate
                  '  Exit Do
               ' End If
                If v = "MOTOS" Then
                    vgTIPODEVEHICULO = 5
                    vgidCampana = 510
                    vgidCia = vgidCia
                ElseIf v = "AUTOS" Then
                    vgTIPODEVEHICULO = 1
                    vgidCampana = 734
                    vgidCia = vgidCia
                ElseIf v = "CAMIONES" Then
                    vgTIPODEVEHICULO = 4
                    vgidCampana = 931
                    vgidCia = vgidCia
                End If

            Case UCase("idasistenciacertificado")
                vCertificado = v
                vgNROPOLIZA = v
            Case UCase("localidad")
                vgLOCALIDAD = Mid(v, 1, 70)
            Case UCase("provincia")
                vgPROVINCIA = Mid(v, 1, 50)
            Case UCase("fechaemision")
                If IsDate(v) Then
                    vgFECHAVIGENCIA = v
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If
            Case UCase("email")
                vgEmail = v
            Case UCase("apellido")
            Case UCase("nombre")
            Case UCase("documento")
                vDocumento = v
                If Len(v) < 16 Then
                    vgNumeroDeDocumento = v
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If
            Case UCase("idtipodoc")
                vgTipodeDocumento = v
'            Case UCase("tipodedocumento")
'                vgTipodeDocumento = v
            Case UCase("fechadenacimiento")
                If IsDate(v) Then
                    vgFechaDeNacimiento = v
               ' Else
                 '   vError = 1
                '    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If
            Case UCase("cuota")
            Case UCase("vigenciaini")
            Case UCase("vigenciafin")
                If IsDate(v) Then
                    vgFECHAVENCIMIENTO = v
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If

            Case UCase("anulado")
            Case UCase("observaciones")
                vOBSERVACIONES = Mid(v, 1, 64)
            Case UCase("calle")
                vgDOMICILIO = Mid(v, 1, 80)
            Case UCase("altura")
                vgDOMICILIO = Mid(vgDOMICILIO & " " & v, 1, 80)
            Case UCase("piso")
                vgDOMICILIO = Mid(vgDOMICILIO & " Piso " & v, 1, 80)
            Case UCase("barrio")
                vgDOMICILIO = Mid(vgDOMICILIO & " " & v, 1, 80)
            Case UCase("partido")
                vgDOMICILIO = Mid(vgDOMICILIO & " " & v, 1, 80)
            Case UCase("cp4")
                If IsNumeric(v) And Len(v) < 5 Then
                    vgCODIGOPOSTAL = v
                End If
            Case UCase("cpa")
            Case UCase("bloquemanzana")
            Case UCase("telparticular")
            Case UCase("telcelular")
                vgTelefono = Mid(v, 1, 10)
            Case UCase("observacionesdomicilio")
            Case UCase("depto")
            Case UCase("razonsocial")
            Case UCase("contacto")
            Case UCase("cuitcuil")
            Case UCase("sexo")
            Case UCase("idestadocivil")
            Case UCase("estadocivil")
            Case UCase("idcondiva")
            Case UCase("condiva")
            Case UCase("nacionalidad")
            Case UCase("entrecalle1")
            Case UCase("entrecalle2")
            Case UCase("idmediodepago")
            Case UCase("mediodepago")
            Case UCase("anio")
                If IsNumeric(v) And v > 1900 And v < Year(Now()) Then
                    vgAno = v
                End If
            Case UCase("patente")
                vPatente = v
                If Len(v) < 10 Then
                    vgPATENTE = v
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If
                
            Case UCase("nromotor")
                vgNroMotor = Mid(v, 1, 50)
            Case UCase("nrocuadro")
            Case UCase("color")
                vgCOLOR = Mid(v, 1, 30)
            Case UCase("empresaapellidoynombre")
                vgAPELLIDOYNOMBRE = Mid(v, 1, 100)
            Case UCase("marca")
                vgMARCADEVEHICULO = Mid(v, 1, 50)
            Case UCase("modelo")
                vgMODELO = Mid(v, 1, 20)
            Case UCase("cilindrada")
            Case UCase("fechahoraemision")
            Case UCase("idproductoasistencia")
                If InStr(vProductosPosibles, Trim(v)) > 0 And IsEmpty(v) = False Then
                v = Mid("0" & v, Len("0" & v) - 1, 2)
                    vgCOBERTURAVEHICULO = v
                    vgCOBERTURAVIAJERO = v
                    vgCOBERTURAHOGAR = v
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & " ; " & "ERROR EN LINEA " & lRow & " VALOR " & v & " EN LA COLUMNA " & sName
                End If
            Case UCase("tarifaasistencia")
            Case UCase("cuponeraasistencia")
            Case UCase("pais")
                vgPais = Mid(v, 1, 30)
            Case UCase("idformadepago")
            Case UCase("formadepago")
            Case UCase("tipodevehiculo")
            Case UCase("productoasistencia")
            Case UCase("insertdate")
        End Select
        Next


    
    
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
 Dim vcamp As Integer
 Dim vdif As Long
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where nroPoliza = '" & Trim(vgNROPOLIZA) & "' and IdCampana = " & vgidCampana
    rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    
    
    
    
    
    
    vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
    vgIDPOLIZA = 0
            If Not rscn1.EOF Then
                vdif = 0  'setea la variale de control de repetido con modificacion en cero
                If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("IdCampana")) <> Trim(vgidCampana) Then vdif = vdif + 1
                If Trim(rscn1("idcia")) <> Trim(vgidCia) Then vdif = vdif + 1
                If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                If Trim(rscn1("LOCALIDAD")) <> Trim(vgLOCALIDAD) Then vdif = vdif + 1
                If Trim(rscn1("PROVINCIA")) <> Trim(vgPROVINCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("Email")) <> Trim(vgEmail) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("IdTipoDoc")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If Trim(rscn1("OBSERVACIONES")) <> Trim(vOBSERVACIONES) Then vdif = vdif + 1
                If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                If Trim(rscn1("PAIS")) <> Trim(vgPais) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If

        rscn1.Close
'-=================================================================================================================
  
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatosAuxilio24 ("
        ssql = ssql & "TIPODEVEHICULO, "
        ssql = ssql & "IdCampana, "
        ssql = ssql & "idcia, "
        ssql = ssql & "NROPOLIZA, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "Email, "
        ssql = ssql & "NumeroDeDocumento, "
        ssql = ssql & "IdTipoDoc, "
        'ssql = ssql & "TipodeDocumento, "
        ssql = ssql & "FechadeNacimiento, "
        ssql = ssql & "FECHAVENCIMIENTO, "
        ssql = ssql & "OBSERVACIONES, "
        ssql = ssql & "DOMICILIO, "
        ssql = ssql & "CODIGOPOSTAL, "
        ssql = ssql & "Telefono, "
        ssql = ssql & "ANO, "
        ssql = ssql & "PATENTE, "
        ssql = ssql & "NroMotor, "
        ssql = ssql & "COLOR, "
        ssql = ssql & "APELLIDOYNOMBRE, "
        ssql = ssql & "MARCADEVEHICULO, "
        ssql = ssql & "MODELO, "
        ssql = ssql & "COBERTURAVEHICULO, "
        ssql = ssql & "COBERTURAVIAJERO, "
        ssql = ssql & "COBERTURAHOGAR, "
        ssql = ssql & "PAIS, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "Modificaciones)"
        ssql = ssql & " values("
        ssql = ssql & Trim(vgTIPODEVEHICULO) & ", "
        ssql = ssql & Trim(vgidCampana) & ", "
        ssql = ssql & Trim(vgidCia) & ", "
        ssql = ssql & Trim(vgNROPOLIZA) & ", '"
        ssql = ssql & Trim(vgLOCALIDAD) & "', '"
        ssql = ssql & Trim(vgPROVINCIA) & "', '"
        ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
        ssql = ssql & Trim(vgEmail) & "', '"
        ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
        'ssql = ssql & Trim(vgTipodeDocumento) & "', '"
        ssql = ssql & Trim(vgTipodeDocumento) & "', '"
        ssql = ssql & Trim(vgFechaDeNacimiento) & "', '"
        ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
        ssql = ssql & Trim(vOBSERVACIONES) & "', '"
        ssql = ssql & Trim(vgDOMICILIO) & "', '"
        ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
        ssql = ssql & Trim(vgTelefono) & "', '"
        ssql = ssql & Trim(vgAno) & "', '"
        ssql = ssql & Trim(vgPATENTE) & "', '"
        ssql = ssql & Trim(vgNroMotor) & "', '"
        ssql = ssql & Trim(vgCOLOR) & "', '"
        ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
        ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
        ssql = ssql & Trim(vgMODELO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
        ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
        ssql = ssql & Trim(vgPais) & "', '"
        ssql = ssql & Trim(vLote) & "', '"
        ssql = ssql & Trim(vdif) & "') "
        cn.Execute ssql
        
        Ll = Ll + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            ll100 = 0
        End If
        DoEvents
    Loop
    oExcel.Workbooks.Close
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
        cn1.CommandTimeout = 600
        cn1.Execute sSPImportacion & " " & lLote & ", " & vUltimaCorrida & ", " & vgidCia & ", " & vgidCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        If rsCMP("UltimaCorridaError") <> "OK" Then
        
            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
            lLote = vLote + 1 'para salir del FOR
        Else
                ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
        End If
        rsCMP.Close
        
    Next lLote
    cn1.Execute "TM_BajaDePolizasHogarYAutos" & " " & vUltimaCorrida & ", " & vgidCia & ", " & vgidCampana & ", " & vTipodeServicio
Exit Sub
errores:
    oExcel.Workbooks.Close
    vgErrores = 1
    If Ll = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & Ll & " Campo: " & vCampo & " Posicion= " & vPosicion
    End If

End Sub



