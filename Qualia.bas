Attribute VB_Name = "Qualia"
Option Explicit

Public Sub ImportarExelQualia()

Dim rscn1 As New Recordset
Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim ssqlProv As String
Dim rsCMPProv As New Recordset
Dim rsp As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v As String, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vgidcia As Long
Dim vgidCampana As Long
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
'Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim sArchivo As String
Dim rsCMP As New Recordset
Dim vHoja As String
Dim lRowH As Long
Dim vtotalDeLineas As Long
Dim vdir As Integer
Dim regMod As Long
'Dim vLote As Long

'Variables propias de Qualia
'
'Dim vgNROPOLIZA As Long
'Variables propias de Qualia
'Dim vfIdDatosAdicionalesQualia As Long
'Dim vfIdPoliza As Long
'Dim vfCodigodeRamo As String
'Dim vfNroTarjetaDeCredito As String
'Dim vfFechaModificacion As Date
Dim vfMotivoBaja As String
Dim vfDescripcionDelProducto As String
Dim vfMontoAsegurado As String
Dim vfFormaDePago As String
'Dim vfModulo As Long
'Dim vfDescripcionDelModulo As String
Dim vfNacionalidad As String
Dim vfTipoTelefono1 As String
Dim vfTipoTelefono2 As String
Dim vfTelefono2 As String
Dim vfSucursalCuenta As String
'Dim vfDigitoVerificador As long
Dim vfPrima As String
Dim vfPremio As String
Dim vfCalleDelBienAsegurado As String
Dim vfNroDelBienAsegurado As String
Dim vfDptoDelBienAsegurado As String
Dim vfPisoDelBienAsegurado As String
Dim vfProvinciaDelBienAsegurado As String
Dim vfCPDelBienAsegurado As String
Dim vfLocalidadDelBienAsegurado As String
Dim vfTecho As String
Dim vfVentanas As String
Dim vfAntiguedad As String
Dim vfTipoCerradura As String
Dim vfDependencias As String
Dim vfAdicionalSeguroTecnico As String
Dim vfAdicionalPalosDeGolf As String
Dim vfAdicionalNotebook As String
Dim vfAdicionalLCD As String
Dim vfAdicionalConsola As String
Dim vfAdicionalConsolaOp1 As String
Dim vfAdicionalConsolaOp2 As String
Dim vfBeneficiario1Nombre As String
Dim vfBeneficiario1Apellido As String
Dim vfBeneficiario1IdTipoDoc As String
Dim vfBeneficiario1NroDocumento As String
Dim vfBeneficiario1Porcentaje As String
Dim vfBeneficiario1Mail As String
Dim vfBeneficiario2Nombre As String
Dim vfBeneficiario2Apellido As String
Dim vfBeneficiario2IdTipoDoc As String
Dim vfBeneficiario2NroDocumento As String
Dim vfBeneficiario2Porcentaje As String
Dim vfBeneficiario2Mail As String
Dim vfBeneficiario3Nombre As String
Dim vfBeneficiario3Apellido As String
Dim vfBeneficiario3IdTipoDoc As String
Dim vfBeneficiario3NroDocumento As String
Dim vfBeneficiario3Porcentaje As String
Dim vfBeneficiario3Mail As String
Dim vfIdEnCliente As String
'Dim vfCanal As Long
'Dim vfFechaDeAlta As Date
'Dim vfFechaDeBaja As String ' valores vacios si no estan dados de baja
'Dim vfCondicionIVA As Long
'Dim vfActividad As Long
Dim vfCodigoDeProducto As String
'Dim vfDescripcionDelVehiculo As String
'Dim vfNumeroDeMotor As String
'Dim VfNumeroDeChasis As String
'Dim vfPatenteDelVehiculo As String
'Dim vfAñoDelVehiculo As Long
'Dim vfValorDelVehiculo As Long
'Dim vfValorDeAccesorios As Long
'Dim vfLugarDeNacimiento As String ' eb la base de prueba esta como string, antes estaba como long
'Dim vfCalle As String 'estaba como long se pasa a asitring
'Dim vfNumero As String
Dim vfDepartamento As String ' estaba como long se pasa a string ( como viene en la base de prueba)
Dim vfPiso As String
'Dim vfCuentaDeDebito As Long
'Dim vfNumeroDelUltimoRecibo As String
'Dim vfDeuda As Long
'Dim vfMarcaCilindroGNC As Long
'Dim vfNumeroDelregulador As Long
'Dim vfNumeroDeCilindroGNC As Long
'Dim vfMarca2CilindroGNC As Long
'Dim vfNumeroDe2CilindroGNC As Long
'Dim vfBeneficiario1TipoDeDocumento As String
'Dim vfBeneficiario1NumeroDeDocumento As String
'Dim vfBeneficiario2TipoDeDocumento As String
'Dim vfBeneficiario2NumeroDeDocumento As String
'Dim vfBeneficiario3TipoDeDocumento As String
'Dim vfNumeroDeCertificado As String
'Dim vfBeneficiario3NumeroDeDocumento As String
Dim vfNroDeTarjetaDeCredito As Long
Dim vfNumeroDelBienAsegurado As String
'Dim vfBANCO As String

Dim vCantDeErrores As Integer
'Dim nroLinea As Long
'Dim lLote As Long

Dim FechaInicial As Date
' -----------------------------------------

On Error Resume Next


    
 
Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
Dim camposParaValidar(106)
'configura las solapas del archivo, agregar o borrar aqui las solapas nuevas
Dim solapas As String
solapas = "BER&BERSA&BSC&BSF&BSJ&"

Dim sFile As New FileSystemObject
Dim fln As TextStream
Set fln = sFile.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
fln.WriteLine "Errores"

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

FechaInicial = Now()

Set oExcel = New Excel.Application
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)

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
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
'=================================================================================================

vHoja = ""
lRowH = 0
vtotalDeLineas = 0
vgidcia = lIdCia


'Comprobación de nombres de las solapas
Dim j As Integer
Dim i As Integer
For i = 1 To 4

  Set oSheet = oBook.Worksheets.Item(i)
  If InStr(1, solapas, oSheet.Name + "&") = 0 Then
        MsgBox "Error en la nombre de la Hoja " & i
        Exit Sub
  ElseIf oSheet.Name = "BSF" Then
       vgidCampana = 934  'SantaFe
  ElseIf oSheet.Name = "BSJ" Then
        vgidCampana = 935 'SanJuan
  ElseIf oSheet.Name = "BERSA" Then
        vgidCampana = 936 'EntreRios
  ElseIf oSheet.Name = "BER" Then
        vgidCampana = 936 'EntreRios
  ElseIf oSheet.Name = "BSC" Then
        vgidCampana = 937 'SantaCruz
  End If
  
Next


cn.Execute "DELETE FROM bandejadeentrada.dbo.importaDatosV2Qualia where idcia=" & 10002043
'rsc.Open "SELECT * FROM bandejadeentrada.dbo.importaDatosV2Qualia where idcia=" & 10002043, cn, adOpenKeyset, adLockOptimistic

If Err.Number Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If


'Cambio de pestaña de excel
For j = 1 To 4

    Set oSheet = oBook.Worksheets(j)
    
    If oSheet.Name = "BSF" Then
       vgidCampana = 934  'SantaFe
    ElseIf oSheet.Name = "BSJ" Then
        vgidCampana = 935 'SanJuan
    ElseIf oSheet.Name = "BERSA" Then
        vgidCampana = 936 'EntreRios
    ElseIf oSheet.Name = "BER" Then
        vgidCampana = 936 'EntreRios
    ElseIf oSheet.Name = "BSC" Then
        vgidCampana = 937 'SantaCruz
    End If


      columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
      extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

      'columnas = extremos(0)
      filas = extremos(1)

camposParaValidar(0) = "Número de Póliza"
camposParaValidar(1) = "Número de Certificado"
camposParaValidar(2) = "Código de Ramo"
camposParaValidar(3) = "Descripción del Ramo"
camposParaValidar(4) = "Sucursal"
camposParaValidar(5) = "Promotor / Vendedor"
camposParaValidar(6) = "Canal"
camposParaValidar(7) = "Nombre del Cliente"
camposParaValidar(8) = "Apellido del cliente"
camposParaValidar(9) = "ID Cliente"
camposParaValidar(10) = "Tipo de Paquete"
camposParaValidar(11) = "Numero de Paquete"
camposParaValidar(12) = "Tipo de Documento"
camposParaValidar(13) = "Número de Documento"
camposParaValidar(14) = "Sexo"
camposParaValidar(15) = "Cuit / Cuil Cliente"
camposParaValidar(16) = "Número de Tarjeta de crédito"
camposParaValidar(17) = "CBU"
camposParaValidar(18) = "Sucursal Cuenta"
camposParaValidar(19) = "Fecha de Alta"
camposParaValidar(20) = "Fecha de Ultima modificación"
camposParaValidar(21) = "Fecha Baja"
camposParaValidar(22) = "Motivo Baja"
camposParaValidar(23) = "Condición Iva"
camposParaValidar(24) = "Actividad"
camposParaValidar(25) = "Descripción del producto"
camposParaValidar(26) = "Código de producto"
camposParaValidar(27) = "Monto Asegurado"
camposParaValidar(28) = "Forma de pago"
camposParaValidar(29) = "Descripción del vehículo"
camposParaValidar(30) = "Módulo"
camposParaValidar(31) = "Descripción del modulo"
camposParaValidar(32) = "Número de Motor"
camposParaValidar(33) = "Número de Chasis"
camposParaValidar(34) = "Patente del vehículo"
camposParaValidar(35) = "Año del vehículo"
camposParaValidar(36) = "Valor del vehículo"
camposParaValidar(37) = "Sin Uso"
camposParaValidar(38) = "Valor de accesorios"
camposParaValidar(39) = "Sin Uso"
camposParaValidar(40) = "Sin Uso"
camposParaValidar(41) = "Lugar de Nacimiento"
camposParaValidar(42) = "Nacionalidad"
camposParaValidar(43) = "País"
camposParaValidar(44) = "Código Postal"
camposParaValidar(45) = "Calle"
camposParaValidar(46) = "Numero"
camposParaValidar(47) = "Departamento"
camposParaValidar(48) = "Piso"
camposParaValidar(49) = "Localidad"
camposParaValidar(50) = "Provincia"
camposParaValidar(51) = "Mail Cliente"
camposParaValidar(52) = "Tipo telefono 1"
camposParaValidar(53) = "Telefono 1"
camposParaValidar(54) = "Tipo telefono 2"
camposParaValidar(55) = "Telefono 2"
camposParaValidar(56) = "Cuenta de Débito"
camposParaValidar(57) = "Sucursal Cuenta"
camposParaValidar(58) = "Digito verificador"
camposParaValidar(59) = "Número del ultimo recibo"
camposParaValidar(60) = "Prima"
camposParaValidar(61) = "Premio"
camposParaValidar(62) = "Deuda"
camposParaValidar(63) = "Calle"
camposParaValidar(64) = "Numero"
camposParaValidar(65) = "Departamento"
camposParaValidar(66) = "Piso"
camposParaValidar(67) = "Provincia"
camposParaValidar(68) = "Código Postal"
camposParaValidar(69) = "Localidad"
camposParaValidar(70) = "Techo"
camposParaValidar(71) = "Ventanas"
camposParaValidar(72) = "Antigüedad años"
camposParaValidar(73) = "Tipo Cerradura"
camposParaValidar(74) = "Dependencias"
camposParaValidar(75) = "Adicional Seguro Técnico (S/N)"
camposParaValidar(76) = "Adicional Palos de Golf (S/N)"
camposParaValidar(77) = "Adicional Notebook (S/N)"
camposParaValidar(78) = "Adicional LCD (S/N)"
camposParaValidar(79) = "Adicional consola (S/N)"
camposParaValidar(80) = "Adicional consola opsción 1(S/N)"
camposParaValidar(81) = "Adicional consola opsción 2(S/N)"
camposParaValidar(82) = "Sin Uso"
camposParaValidar(83) = "Marca Cilindro GNC"
camposParaValidar(84) = "Número de Regulador"
camposParaValidar(85) = "Marca Cilindro GNC"
camposParaValidar(86) = "Número de Cilindro GNC"
camposParaValidar(87) = "Marca 2° Cilindro GNC"
camposParaValidar(88) = "Número de 2° Cilindro GNC"
camposParaValidar(89) = "Beneficiario1 Nombre"
camposParaValidar(90) = "Beneficiario1 Apellido"
camposParaValidar(91) = "Beneficiario1  Tipo de Documento"
camposParaValidar(92) = "Beneficiario1 Número de Documento"
camposParaValidar(93) = "Beneficiario1 Porcentaje"
camposParaValidar(94) = "Beneficiario1 Mail"
camposParaValidar(95) = "Beneficiario2 Nombre"
camposParaValidar(96) = "Beneficiario2 Apellido"
camposParaValidar(97) = "Beneficiario2  Tipo de Documento"
camposParaValidar(98) = "Beneficiario2 Número de Documento"
camposParaValidar(99) = "Beneficiario2 Porcentaje"
camposParaValidar(100) = "Beneficiario2 Mail"
camposParaValidar(101) = "Beneficiario3 Nombre"
camposParaValidar(102) = "Beneficiario3 Apellido"
camposParaValidar(103) = "Beneficiario3  Tipo de Documento"
camposParaValidar(104) = "Beneficiario3 Número de Documento"
camposParaValidar(105) = "Beneficiario3 Porcentaje"
camposParaValidar(106) = "Beneficiario3 Mail"

      If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then
      
        lRow = 2
          col.RemoveAll
          lCol = 1
          Do While lCol < columnas + 1
              v = oSheet.Cells(1, lCol)
              If IsEmpty(v) Then Exit Do
              sName = v
              col.Add lCol, v
              lCol = lCol + 1
          Loop
          vdir = 0
          
          Do While lRow < filas + 1
          
            '====maneja los lotes para corte de importacion========
            nroLinea = nroLinea + 1
            If nroLinea = LongDeLote + 1 Then
                vLote = vLote + 1
                nroLinea = 1
            End If
            '======================================================

'              rsc.AddNew
'            vCantDeErrores = 0
'                rsc("idcia") = vgidcia  '  10002043
'                rsc("idcampana") = vgidCampana
'                rsc("IdLote") = vLote

              For lCol = 1 To columnas
                  sName = col.Item(lCol)
                  v = ""
                  v = Trim(oSheet.Cells(lRow, lCol))
                If IsEmpty(v) = False Then
              
                  If lCol = 1 And IsEmpty(v) Then
                      Exit Do
                  End If
                  
                    'sName = col.Item(lCol)
                  
                  Select Case UCase(Trim(sName))
                        Case UCase("Número de póliza")
'                            rsc("Nropoliza").Value = v
                            vgNROPOLIZA = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Número de certificado")
'                            rsc("Nrosecuencial").Value = Trim(v)
                            vgNROSECUENCIAL = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Código de ramo")
'                            rsc("CodigoDeRamo") = v
                            vgCodigoDeRamo = v
                            ssql = "select idproductomultiasistencias from tm_productosmultiasistencias where idcampana= " & vgidCampana & " and idProductoEnCliente = '" & v & "'"
                            rsp.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                            
                            If Not rsp.EOF Then
                                vgIdProducto = rsp("idproductomultiasistencias")
                            Else
                                vgIdProducto = -99
                            End If
                            
                            rsp.Close
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                               
                        Case UCase("Descripción del ramo")
'                            rsc("Ocupacion").Value = v
'                            vgRama = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
'                        Case UCase("Sucursal")
'                           vfagencia = v
'                        Case Ucase("Promotor/Vendedor"
'                          rsc("Cargo") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Canal")
'                           rsc("Canal") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Nombre del cliente")
'                           rsc("Nombre") = v
                           vgNombre = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Apellido del cliente")
'                           rsc("Apellido") = v
                           vgApellido = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("ID Cliente")
'                           rsc("IdEnCliente") = v
                            vfIdEnCliente = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Tipo de paquete")
'                           rsc("TipoDeServicio") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Número de paquete")
'                           rsc("IdTiopoDePoliza") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Tipo de documento")
'                            rsc("TipoDeDocumento") = v
                            vgTipodeDocumento = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Número de documento")
'                           rsc("NumerodeDocumento") = v
                           vgNumeroDeDocumento = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        Case UCase("Sexo")
'                            rsc("Sexo") = v
                            vgSexo = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Cuit / Cuil cliente")
'                           rsc("Cuit") = v
                           vgCuit = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Número de Tarjet de Crédito")
'                            rsc("NroTarjetaDeCredito") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("CBU"
'                           rsc("CBU") = v
'                        Case Ucase("Sucursal cuenta"
'                            rsc("SucursalCuenta") = v
                        Case UCase("Fecha de alta")
'                           If IsNumeric(v) Then rsc("FechaVigencia") = AAMMDDToDD_MM_AA(v)
                            If IsNumeric(v) Then vgFECHAVIGENCIA = AAMMDDToDD_MM_AA(v)
                            If IsNumeric(v) Then vgFECHAALTAOMNIA = AAMMDDToDD_MM_AA(v)
'                           vgFECHAVIGENCIA = v
'                           vgFECHAALTAOMNIA = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Fecha de última modificación")
'                            If IsNumeric(v) Then rsc("FechaModificación") = AAMMDDToDD_MM_AA(v)
                            If IsNumeric(v) Then vgFechaModificacion = AAMMDDToDD_MM_AA(v)
'                            vgFechaModificacion = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Fecha baja")
                        
                        
                           If IsNumeric(v) Then
'                            rsc("FechaVencimiento") = AAMMDDToDD_MM_AA(v)
'                            rsc("fechabajaomnia") = AAMMDDToDD_MM_AA(v)
                            If IsNumeric(v) Then vgFECHABAJAOMNIA = AAMMDDToDD_MM_AA(v)
                            If IsNumeric(v) Then vgFECHAVENCIMIENTO = AAMMDDToDD_MM_AA(v)
'                            vgFECHABAJAOMNIA = v
'                            vgFECHAVENCIMIENTO = v
                            
                           End If
                           
                           
                           
'                           If IsNull(v) Then
'
'                            vgFECHAVENCIMIENTO = Null
'                            vgFECHABAJAOMNIA = Null
'                            vgFECHABAJAOMNIA = v
'                           End If
                           
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Motivo baja")
'                            rsc("MotivoBaja") = v
'                            vfMotivoBaja = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("Condición IVA"
'                           rsc("CondicionIVA") = v
                        Case UCase("Actividad")
'                            rsc("Operacion") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
'
                        Case UCase("Descripción del producto")
'                           rsc("DescripcionDelProducto") = v
                            vfDescripcionDelProducto = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Código de producto")
'                            rsc("CodigoDeProducto") = v
                            vfCodigoDeProducto = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Monto asegurado")
'                           rsc("MontoAsegurado") = v
                            vfMontoAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Forma de pago")
'                            rsc("FormaDePago") = v
                                vfFormaDePago = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("Descripción del vehículo"
'                           rsc("Descripción del vehículo") = v
'                        Case UCase("Módulo")
'                            rsc("Modulo") = v
'                            vfmodulo = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case UCase("Descripción del modulo")
'                           rsc("DescripcionDelModulo") = Trim(v)
'                           vgDescripcionDelModulo = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
'
'                        Case Ucase("Número de motor"
'                            rsc("Número de motor") = v
'                        Case Ucase("Número de Chasis"
'                           rsc("Número de Chasis") = v
'                        Case Ucase("Patente del vehículo"
'                            rsc("TPatente del vehículo") = v
'                        Case Ucase("Año del vehículo"
'                           rsc("Año del vehículo") = v
'                        Case Ucase("Valor del vehículo"
'                            rsc("Valor del vehículo") = v
'                        Case Ucase("Valor de accesorios"
'                          rsc("ValorDeAccesorios") = v
                        Case UCase("Lugar de nacimiento")
'                            rsc("PaisOrigen") = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Nacionalidad")
'                           rsc("Nacionalidad") = v
                            vfNacionalidad = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
                        
                        Case UCase("País")
                            If vdir = 0 Then vgPais = v

                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Código postal")
                           If vdir = 0 Then vgCODIGOPOSTAL = v
                           If vdir = 1 Then vfCPDelBienAsegurado = v
                           
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Calle")
                            If vdir = 0 Then vgDOMICILIO = v
'                            vgDOMICILIO = v
                            If vdir = 1 Then vfCalleDelBienAsegurado = v
'                            vgDOMICILIO = v
                            
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Numero")
                        
                           If vdir = 0 And Len(vgDOMICILIO) > 0 Then vgDOMICILIO = vgDOMICILIO & " " & v
                           If IsNumeric(v) And vdir = 1 Then vfNroDelBienAsegurado = v
                           
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Departamento")
                            If vdir = 0 And Len(vgDOMICILIO) > 0 Then vgDOMICILIO = vgDOMICILIO & " " & v
                            If vdir = 1 Then vfDptoDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                            
                        Case UCase("Piso")
                           If vdir = 0 And Len(vgDOMICILIO) > 0 Then vgDOMICILIO = vgDOMICILIO & " " & v
                           If vdir = 1 Then vfPisoDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Localidad")
                            If vdir = 0 Then vgLOCALIDAD = v
'                            vgLOCALIDAD = v
                            If vdir = 1 Then vfLocalidadDelBienAsegurado = v
'                            vgLOCALIDAD = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Provincia")
'                            If v <> 0 Then
                                ssqlProv = "Select denominacion from dicprovincias where idprovinciaqualia=" & v
                                rsCMPProv.Open ssqlProv, cn1, adOpenForwardOnly, adLockReadOnly
                             
                                If Not rsCMPProv.EOF Then
                                    v = rsCMPProv("denominacion")
                                    If vdir = 0 Then vgPROVINCIA = v
                                    If vdir = 1 Then vfProvinciaDelBienAsegurado = v
                                    vdir = 1
                                End If
                                rsCMPProv.Close
'                            End If
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Mail cliente")
                            vgEmail = v
                            vgEmail = v
                            
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Tipo telefono 1")
                           vfTipoTelefono1 = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Telefono 1")
'                            rsc("Telefono") = Trim(v)
                            vgTelefono = Trim(v)
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Tipo telefono 2")
                           vfTipoTelefono2 = v
                           
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Telefono 2")
'                            rsc("Telefono2") = Trim(v)
                            vfTelefono2 = Trim(v)
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("Cuenta de débito"
'                          rsc("CuentaDeDebito") = v
'                        Case UCase("Sucursal cuenta")
'                            vfSucursalCuenta = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)
'
'                        Case UCase("Digito verificador") ' no esta definido como variable global o variable formulario
'                           vgdigitoVerificador = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("Número del ultimo recibo"
'                            rsc("NroUltimoRecibo") = v
                        Case UCase("Prima")
                           vfPrima = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Premio")
                            vfPremio = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case UCase("Deuda") ' no esta definido como variable global o variable formulario
'                           vfdeuda = v
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Calle")
                            If vdir = 1 Then vfCalleDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Numero")
                           If IsNumeric(v) And vdir = 1 Then vfNumeroDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Departamento")
                            If vdir = 1 Then vfDptoDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Piso")
                           If vdir = 1 Then vfPisoDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Provincia")
                            If vdir = 1 Then vfProvinciaDelBienAsegurado = v
                            vgPROVINCIA = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Código postal")
                           If vdir = 1 Then vfCPDelBienAsegurado = v
                            vgCODIGOPOSTAL = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Localidad")
                            If vdir = 1 Then vfLocalidadDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Techo")
                           vfTecho = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Ventanas")
                            vfVentanas = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Antigüedad años")
                            vfAntiguedad = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Tipo cerradura")
                            vfTipoCerradura = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Dependencias")
                            vfDependencias = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional Seguro Técnico (S/N)")
                            vfAdicionalSeguroTecnico = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional Palos de Golf (S/N)")
                            vfAdicionalPalosDeGolf = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional Notebook (S/N)")
                            vfAdicionalNotebook = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional LCD (S/N)")
                            vfAdicionalLCD = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional consola (S/N)")
                            vfAdicionalConsola = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional consola opción 1(S/N)")
                            vfAdicionalConsolaOp1 = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Adicional consola opsción 2 (S/N)")
                            vfAdicionalConsolaOp2 = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

'                        Case Ucase("Marca cilindro GNC"
'                            rsc("Marca cilindro GNC") = v
'                        Case Ucase("Número de regulador"
'                            rsc("Número de regulador") = v
'                        Case Ucase("Marca cilindro GNC"
'                            rsc("Marca cilindro GNC") = v
'                        Case Ucase("Número de cilindro GNC"
'                            rsc("Número de cilindro GNC") = v
'                        Case Ucase("Marca 2º cilindro GNC"
'                            rsc("Marca 2º cilindro GNC") = v
'                        Case Ucase("Número de 2º cilindro GNC"
'                            rsc("Número de 2º cilindro GNC") = v
                        Case UCase("Beneficiario1 Nombre")
                            vfBeneficiario1Nombre = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario1 Apellido")
                            vfBeneficiario1Apellido = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario1 Tipo de documento")
                            vfBeneficiario1IdTipoDoc = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario1 Número de documento")
                            vfBeneficiario1NroDocumento = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario1 Porcentaje")
                            vfBeneficiario1Porcentaje = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario1 Mail")
                            vfBeneficiario1Mail = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Nombre")
                            vfBeneficiario2Nombre = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Apellido")
                            vfBeneficiario2Apellido = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Tipo de documento")
'                            rsc("Beneficiario2Beneficiario2IdTipoDoc") = v
                            vfNroDelBienAsegurado = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Número de documento")
'                            rsc("Beneficiario2NroDocumento") = v
                                vfBeneficiario2NroDocumento = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Porcentaje")
'                            rsc("Beneficiario2Porcentaje") = v
                            vfBeneficiario2Porcentaje = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario2 Mail")
 '                           rsc("Beneficiario2Mail") = v
                            vfBeneficiario2Mail = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Nombre")
  '                          rsc("Beneficiario3Nombre") = v
                            vfBeneficiario3Nombre = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Apellido")
   '                         rsc("Beneficiario3Apellido") = v
                            vfBeneficiario3Apellido = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Tipo de documento")
    '                        rsc("Beneficiario3IdTipoDoc") = v
                            vfBeneficiario3IdTipoDoc = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Número de documento")
     '                       rsc("Beneficiario3NroDocumento") = v
                            vfBeneficiario3NroDocumento = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Porcentaje")
      '                      rsc("Beneficiario3Porcentaje") = v
                            vfBeneficiario3Porcentaje = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                        Case UCase("Beneficiario3 Mail")
       '                     rsc("Beneficiario3Mail") = v
                            vfBeneficiario3Mail = v
                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, oSheet.Name, lRow, sName)

                  End Select
                  
                End If
                
              Next
              
              vgAPELLIDOYNOMBRE = vgApellido & ", " & vgNombre
                
                '=================================================================================
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
' cn1.Execute " TM_NormalizaPolizas " & vgNROPOLIZA & ", " & vNroPolizaB & " OUT"
ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
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
'                If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                'If Trim(rscn1("FECHABAJAOMNIA")) <> Trim(vgFECHABAJAOMNIA) Then vdif = vdif + 1
                'If Trim(rscn1("IDAUTO")) <> Trim(vgIDAUTO) Then vdif = vdif + 1
                'If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                'If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                'If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
                'If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                'If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                'If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                'If Trim(rscn1("TipodeServicio")) <> Trim(vgTipodeServicio) Then vdif = vdif + 1
'                If Trim(rscn1("IDTIPODECOBERTURA")) <> Trim(vgIDTIPODECOBERTURA) Then vdif = vdif + 1
                'If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                'If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                'If Trim(rscn1("TipodeOperacion")) <> Trim(vgTipodeOperacion) Then vdif = vdif + 1
                'If Trim(rscn1("Operacion")) <> Trim(vgOperacion) Then vdif = vdif + 1
                'If Trim(rscn1("CATEGORIA")) <> Trim(vgCATEGORIA) Then vdif = vdif + 1
                'If Trim(rscn1("ASISTENCIAXENFERMEDAD")) <> Trim(vgASISTENCIAXENFERMEDAD) Then vdif = vdif + 1
'                If Trim(rscn1("IdCampana")) <> Trim(vgIdCampana) Then vdif = vdif + 1
                'If Trim(rscn1("Conductor")) <> Trim(vgConductor) Then vdif = vdif + 1
                'If Trim(rscn1("CodigoDeProductor")) <> Trim(vgCodigoDeProductor) Then vdif = vdif + 1
               ' If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("NumeroDeDocumento")) <> Trim(vgNumeroDeDocumento) Then vdif = vdif + 1
                'If Trim(rscn1("TipodeHogar")) <> Trim(vgTipodeHogar) Then vdif = vdif + 1
                'If Trim(rscn1("IniciodeAnualidad")) <> Trim(vgIniciodeAnualidad) Then vdif = vdif + 1
                'If Trim(rscn1("PolizaIniciaAnualidad")) <> Trim(vgPolizaIniciaAnualidad) Then vdif = vdif + 1
                If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                'If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                'If Trim(rscn1("Gama")) <> Trim(vgGama) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
            End If

        rscn1.Close
'-=================================================================================================================
 


'        If DateDiff("d", vgFECHAVENCIMIENTO, Now()) < 0 Then
            ssql = "Insert into bandejadeentrada.dbo.ImportaDatosV2Qualia ("
                ssql = ssql & "IDPOLIZA,"
                ssql = ssql & "IDCIA,"
                ssql = ssql & "NUMEROCOMPANIA,"
                ssql = ssql & "NROPOLIZA,"
                ssql = ssql & "NROSECUENCIAL,"
                ssql = ssql & "APELLIDOYNOMBRE,"
                ssql = ssql & "DOMICILIO,"
                ssql = ssql & "LOCALIDAD,"
                ssql = ssql & "PROVINCIA,"
                ssql = ssql & "CODIGOPOSTAL,"
                ssql = ssql & "FECHAVIGENCIA,"
                ssql = ssql & "FECHAVENCIMIENTO,"
                ssql = ssql & "FECHAALTAOMNIA,"
                ssql = ssql & "FECHABAJAOMNIA,"
                ssql = ssql & "IDAUTO,"
                ssql = ssql & "MARCADEVEHICULO,"
                ssql = ssql & "MODELO,"
                ssql = ssql & "COLOR,"
                ssql = ssql & "ANO,"
                ssql = ssql & "PATENTE,"
                ssql = ssql & "TIPODEVEHICULO,"
                ssql = ssql & "TipodeServicio,"
                ssql = ssql & "IDTIPODECOBERTURA,"
                ssql = ssql & "COBERTURAVEHICULO,"
                ssql = ssql & "COBERTURAVIAJERO,"
                ssql = ssql & "TipodeOperacion,"
                ssql = ssql & "Operacion,"
                ssql = ssql & "CATEGORIA,"
                ssql = ssql & "ASISTENCIAXENFERMEDAD,"
                ssql = ssql & "CORRIDA,"
                ssql = ssql & "FECHACORRIDA,"
                ssql = ssql & "IdCampana,"
                ssql = ssql & "Conductor,"
                ssql = ssql & "CodigoDeProductor,"
                ssql = ssql & "CodigoDeServicioVip,"
                ssql = ssql & "TipodeDocumento,"
                ssql = ssql & "NumeroDeDocumento,"
                ssql = ssql & "TipodeHogar,"
                ssql = ssql & "IniciodeAnualidad,"
                ssql = ssql & "PolizaIniciaAnualidad,"
                ssql = ssql & "Telefono,"
                ssql = ssql & "NroMotor,"
                ssql = ssql & "Gama,"
                ssql = ssql & "InformadoSinCobertura,"
                ssql = ssql & "MontoCoverturaVidrios,"
                ssql = ssql & "COBERTURAHOGAR,"
                ssql = ssql & "CodigoDeProceso,"
                ssql = ssql & "IdTipodePoliza,"
                ssql = ssql & "Referido,"
                ssql = ssql & "Telefono2,"
                ssql = ssql & "Telefono3,"
                ssql = ssql & "IdProducto,"
                ssql = ssql & "email,"
                ssql = ssql & "email2,"
                ssql = ssql & "pais,"
                ssql = ssql & "Sexo,"
                ssql = ssql & "Agencia,"
                ssql = ssql & "Codigoencliente,"
                ssql = ssql & "DocumentoReferente,"
                ssql = ssql & "Ocupacion,"
                ssql = ssql & "Cargo,"
                ssql = ssql & "Nombre,"
                ssql = ssql & "Apellido,"
                ssql = ssql & "IdEnCliente,"
                ssql = ssql & "Cuit,"
                ssql = ssql & "NroDeTarjetaDeCredito,"
                ssql = ssql & "CBU,"
                ssql = ssql & "SucursalCuenta,"
                ssql = ssql & "FechaModificacion,"
                ssql = ssql & "MotivoBaja,"
                ssql = ssql & "CondicionIVA,"
                ssql = ssql & "DescripcionDelProducto,"
                ssql = ssql & "MontoAsegurado,"
                ssql = ssql & "FormaDePago,"
                ssql = ssql & "DescripcionDelMundo,"
                ssql = ssql & "PaisOrigen,"
                ssql = ssql & "Nacionalidad,"
                ssql = ssql & "TipoTelefono1,"
                ssql = ssql & "TipoTelefono2,"
                ssql = ssql & "NroUltimoRecibo,"
                ssql = ssql & "Prima,"
                ssql = ssql & "Premio,"
                ssql = ssql & "CalleDelBienAsegurado,"
                ssql = ssql & "NumeroDelBienAsegurado,"
                ssql = ssql & "DptoDelBienAsegurado,"
                ssql = ssql & "ProvinciaDelBienAsegurado,"
                ssql = ssql & "CPDelBienAsegurado,"
                ssql = ssql & "LocalidadDelBienAsegurado,"
                ssql = ssql & "Techo,"
                ssql = ssql & "Ventanas,"
                ssql = ssql & "Antiguedad,"
                ssql = ssql & "TipoCerradura,"
                ssql = ssql & "Dependencias,"
                ssql = ssql & "AdicionalSeguroTecnico,"
                ssql = ssql & "AdicionalPalosDeGolf,"
                ssql = ssql & "AdicionalNotebook,"
                ssql = ssql & "AdicionalLCD,"
                ssql = ssql & "AdicionalConsolaOp1,"
                ssql = ssql & "AdicionalConsolaOp2,"
                ssql = ssql & "Beneficiario1Nombre,"
                ssql = ssql & "Beneficiario1Apellido,"
                ssql = ssql & "Beneficiario1IdTipoDoc,"
                ssql = ssql & "Beneficiario1NroDocumento,"
                ssql = ssql & "Beneficiario1Porcentaje,"
                ssql = ssql & "Beneficiario1Mail,"
                ssql = ssql & "Beneficiario2Nombre,"
                ssql = ssql & "Beneficiario2Apellido,"
                ssql = ssql & "Beneficiario2IdTipoDoc,"
                ssql = ssql & "Beneficiario2NroDocumento,"
                ssql = ssql & "Beneficiario2Porcentaje,"
                ssql = ssql & "Beneficiario2Mail,"
                ssql = ssql & "Beneficiario3Nombre,"
                ssql = ssql & "Beneficiario3Apellido,"
                ssql = ssql & "Beneficiario3IdTipoDoc,"
                ssql = ssql & "Beneficiario3NroDocumento,"
                ssql = ssql & "Beneficiario3Porcentaje,"
                ssql = ssql & "Beneficiario3Mail,"
                ssql = ssql & "NroDelBienAsegurado,"
                ssql = ssql & "PisoDelBienAsegurado,"
                ssql = ssql & "AdicionalConsola,"
                ssql = ssql & "CodigoDeProducto,"
                ssql = ssql & "IdLote, "
                ssql = ssql & "Modificaciones )"
                ssql = ssql & " values("
                ssql = ssql & Trim(vgIDPOLIZA) & ", "
                ssql = ssql & Trim(vgidcia) & ", '"
                ssql = ssql & Trim(vgNUMEROCOMPANIA) & "', '"
                ssql = ssql & Trim(vgNROPOLIZA) & "', '"
                ssql = ssql & Trim(vgNROSECUENCIAL) & "', '"
                ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
                ssql = ssql & Trim(vgDOMICILIO) & "', '"
                ssql = ssql & Trim(vgLOCALIDAD) & "', '"
                ssql = ssql & Trim(vgPROVINCIA) & "', '"
                ssql = ssql & Trim(vgCODIGOPOSTAL) & "', '"
                ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
                ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
                ssql = ssql & Trim(vgFECHAALTAOMNIA) & "', '"
                ssql = ssql & Trim(vgFECHABAJAOMNIA) & "', '"
                ssql = ssql & Trim(vgIDAUTO) & "', '"
                ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
                ssql = ssql & Trim(vgMODELO) & "', '"
                ssql = ssql & Trim(vgCOLOR) & "', '"
                ssql = ssql & Trim(vgAno) & "', '"
                ssql = ssql & Trim(vgPATENTE) & "', '"
                ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
                ssql = ssql & Trim(vgTipodeServicio) & "', '"
                ssql = ssql & Trim(vgIDTIPODECOBERTURA) & "', '"
                ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
                ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', '"
                ssql = ssql & Trim(vgTipodeOperacion) & "', '"
                ssql = ssql & Trim(vgOperacion) & "', '"
                ssql = ssql & Trim(vgCATEGORIA) & "', '"
                ssql = ssql & Trim(vgASISTENCIAXENFERMEDAD) & "', '"
                ssql = ssql & Trim(vgCORRIDA) & "', '"
                ssql = ssql & Trim(vgFECHACORRIDA) & "', '"
                ssql = ssql & Trim(vgidCampana) & "', '"
                ssql = ssql & Trim(vgConductor) & "', '"
                ssql = ssql & Trim(vgCodigoDeProductor) & "', '"
                ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
                ssql = ssql & Trim(vgTipodeDocumento) & "', '"
                ssql = ssql & Trim(vgNumeroDeDocumento) & "', '"
                ssql = ssql & Trim(vgTipodeHogar) & "', '"
                ssql = ssql & Trim(vgIniciodeAnualidad) & "', '"
                ssql = ssql & Trim(vgPolizaIniciaAnualidad) & "', '"
                ssql = ssql & Trim(vgTelefono) & "', '"
                ssql = ssql & Trim(vgNroMotor) & "', '"
                ssql = ssql & Trim(vgGama) & "', '"
                ssql = ssql & Trim(vgInformadoSinCobertura) & "', '"
                ssql = ssql & Trim(vgMontoCoverturaVidrios) & "', '"
                ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
                ssql = ssql & Trim(vgCodigoDeProceso) & "', '"
                ssql = ssql & Trim(vgIdTipoDePOliza) & "', '"
                ssql = ssql & Trim(vgReferido) & "', '"
                ssql = ssql & Trim(vgTelefono2) & "', '"
                ssql = ssql & Trim(vgTelefono3) & "', '"
                ssql = ssql & Trim(vgIdProducto) & "', '"
                ssql = ssql & Trim(vgEmail) & "', '"
                ssql = ssql & Trim(vgemail2) & "', '"
                ssql = ssql & Trim(vgPais) & "', '"
                ssql = ssql & Trim(vgSexo) & "', '"
                ssql = ssql & Trim(vgAgencia) & "', '"
                ssql = ssql & Trim(vgCodigoEnCliente) & "', '"
                ssql = ssql & Trim(vgDocumentoReferente) & "', '"
                ssql = ssql & Trim(vgOcupacion) & "', '"
                ssql = ssql & Trim(vgCargo) & "', '"
                ssql = ssql & Trim(vgNombre) & "', '"
                ssql = ssql & Trim(vgApellido) & "', '"
                ssql = ssql & Trim(vfIdEnCliente) & "', '"
                ssql = ssql & Trim(vgCuit) & "', '"
                ssql = ssql & Trim(vfNroDeTarjetaDeCredito) & "', '"
                ssql = ssql & Trim(vgCBU) & "', '"
                ssql = ssql & Trim(vgSucursalCuenta) & "', '"
                ssql = ssql & Trim(vgFechaModificacion) & "', '"
                ssql = ssql & Trim(vfMotivoBaja) & "', '"
                ssql = ssql & Trim(vgCondicionIVA) & "', '"
                ssql = ssql & Trim(vfDescripcionDelProducto) & "', '"
                ssql = ssql & Trim(vfMontoAsegurado) & "', '"
                ssql = ssql & Trim(vfFormaDePago) & "', '"
                ssql = ssql & Trim(vgDescripcionDelMundo) & "', '"
                ssql = ssql & Trim(vgPaisOrigen) & "', '"
                ssql = ssql & Trim(vfNacionalidad) & "', '"
                ssql = ssql & Trim(vfTipoTelefono1) & "', '"
                ssql = ssql & Trim(vfTipoTelefono2) & "', '"
                ssql = ssql & Trim(vgNroUltimoRecibo) & "', '"
                ssql = ssql & Trim(vfPrima) & "', '"
                ssql = ssql & Trim(vfPremio) & "', '"
                ssql = ssql & Trim(vfCalleDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfNumeroDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfDptoDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfProvinciaDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfCPDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfLocalidadDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfTecho) & "', '"
                ssql = ssql & Trim(vfVentanas) & "', '"
                ssql = ssql & Trim(vfAntiguedad) & "', '"
                ssql = ssql & Trim(vfTipoCerradura) & "', '"
                ssql = ssql & Trim(vfDependencias) & "', '"
                ssql = ssql & Trim(vfAdicionalSeguroTecnico) & "', '"
                ssql = ssql & Trim(vfAdicionalPalosDeGolf) & "', '"
                ssql = ssql & Trim(vfAdicionalNotebook) & "', '"
                ssql = ssql & Trim(vfAdicionalLCD) & "', '"
                ssql = ssql & Trim(vfAdicionalConsolaOp1) & "', '"
                ssql = ssql & Trim(vfAdicionalConsolaOp2) & "', '"
                ssql = ssql & Trim(vfBeneficiario1Nombre) & "', '"
                ssql = ssql & Trim(vfBeneficiario1Apellido) & "', '"
                ssql = ssql & Trim(vfBeneficiario1IdTipoDoc) & "', '"
                ssql = ssql & Trim(vfBeneficiario1NroDocumento) & "', '"
                ssql = ssql & Trim(vfBeneficiario1Porcentaje) & "', '"
                ssql = ssql & Trim(vfBeneficiario1Mail) & "', '"
                ssql = ssql & Trim(vfBeneficiario2Nombre) & "', '"
                ssql = ssql & Trim(vfBeneficiario2Apellido) & "', '"
                ssql = ssql & Trim(vfBeneficiario2IdTipoDoc) & "', '"
                ssql = ssql & Trim(vfBeneficiario2NroDocumento) & "', '"
                ssql = ssql & Trim(vfBeneficiario2Porcentaje) & "', '"
                ssql = ssql & Trim(vfBeneficiario2Mail) & "', '"
                ssql = ssql & Trim(vfBeneficiario3Nombre) & "', '"
                ssql = ssql & Trim(vfBeneficiario3Apellido) & "', '"
                ssql = ssql & Trim(vfBeneficiario3IdTipoDoc) & "', '"
                ssql = ssql & Trim(vfBeneficiario3NroDocumento) & "', '"
                ssql = ssql & Trim(vfBeneficiario3Porcentaje) & "', '"
                ssql = ssql & Trim(vfBeneficiario3Mail) & "', '"
                ssql = ssql & Trim(vfNroDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfPisoDelBienAsegurado) & "', '"
                ssql = ssql & Trim(vfAdicionalConsola) & "', '"
                ssql = ssql & Trim(vfCodigoDeProducto) & "', '"
                ssql = ssql & Trim(vLote) & "', '"
                ssql = ssql & Trim(vdif) & "') "
                cn.Execute ssql

'========Control de errores=========================================================
                                If Err Then
                                    vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgidCampana, "Proceso", lRow, "")
                                    Err.Clear
                                
                                End If
'===========================================================================================

                
         '=================================================================================
         
         

'            ssql = "Select idpoliza from tm_polizas where idcampana =" & vgidCampana & " and nroPoliza = '" & vgNROPOLIZA & "'"
'            rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'            If Not rs.EOF Then
'                rsc("idpoliza") = rs("idpoliza")
'            Else
'                rsc("idpoliza") = 0
'            End If
'
'            rs.Close
'
'            'Graba si no hay errores
'            If vCantDeErrores = 0 Then
'                rsc.Update
'            End If

              vdir = 0
              lRow = lRow + 1
              lRowH = lRowH + 1
              If vHoja <> oSheet.Name Then
                vHoja = oSheet.Name
                lRowH = 1
              End If
    '=================================================================================

'               If vdif > 0 Then
'            regMod = regMod + 1
'        End If
              ll100 = ll100 + 1
              If ll100 = 1 Then 'cambiar el valor para agrandar el tiempo entre refresco de msj
                  ImportadordePolizas.txtProcesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 1 & " de la Hoja: " & oSheet.Name & " de un total de " & filas

            ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
            If vdif > 0 Then
                regMod = regMod + 1
            End If
            ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & vgidCampana & "and corrida =" & vgCORRIDA
            cn1.Execute ssql
            '=========================================================================================================================================================


                  ll100 = 0
              End If
              DoEvents



          Loop
    
    Else
        MsgBox ("Los siguientes campos obligatorios de la hoja nro " & j & " no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
        oExcel.Workbooks.Close
        Set oExcel = Nothing

        Exit Sub
    End If
    
    vtotalDeLineas = vtotalDeLineas + filas
Next j
    
'================Control de Leidos===============================================
                            cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & vtotalDeLineas
                            listoParaProcesar
'=================================================================================
    
rsc.Close

    
    ImportadordePolizas.txtProcesando.Text = " Procesando un total de " & vtotalDeLineas & " datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos===========================================
                            cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
'    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
'    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
'    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    
    cn1.Execute "update tm_campana set UltimaCorridaError='' , UltimaCorridaCantidadderegistros=0  where idcontacto= " & vgidcia
        ImportadordePolizas.txtProcesando.BackColor = &HC0C0FF
    
    ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    DoEvents
'    cn1.CommandTimeout = 300
'    cn1.Execute sSPImportacion
'
'    ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & vgidcampana
'    rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'    vRegistrosProcesados = vRegistrosProcesados + rsCMP("UltimaCorridaCantidadImportada")
'    If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
'        MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
'    End If
'    rsCMP.Close
    
     For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'        If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
'            MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
'            lLote = vLote + 1 'para salir del FOR
'        Else
                ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
                DoEvents
'        End If

        rsCMP.Close
    Next lLote

    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidcia & ", " & vgidCampana

'============Finaliza Proceso========================================================
                            cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxcia " & lIdCia & ", " & vgCORRIDA
                            Procesado
'=====================================================================================
   
    
        ImportadordePolizas.txtProcesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & vlineasTotales & " FinDeProceso"
        ImportadordePolizas.txtProcesando.BackColor = &HFFFFFF
    
    
    cn1.Execute "update tm_campana set  UltimaCorridaCantidadderegistros = " & vRegistrosProcesados & " where idcontacto=" & vgidcia

oExcel.Workbooks.Close
Set oExcel = Nothing

Exit Sub
errores:
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow & " Columna: " & sName
    End If



End Sub

