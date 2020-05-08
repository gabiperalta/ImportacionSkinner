Attribute VB_Name = "General"
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'Private WithEvents mobjPkgEvents As DTS.Package
Global SCAmino As String
Global cn As New ADODB.Connection
Global cn1 As New ADODB.Connection
Global cn2 As New ADODB.Connection
Global lIdCia As Long, lidCampana As Long, lIdCampanaCall As Long, lIdUtempresaCall As Long
Global fileimportacion As String, sSPImportacion As String
Global sDirImportacion As String, sdtsImportacion As String, sfileimportacion As String
Global vgErrores As Integer

'----Variables De Importacion de Texto-----------------------------------------
Global vgIDPOLIZA As Long
Global vgidCia As Long
Global vgidCampana As Integer
Global vgidUtEmpresa As Integer
Global vgidCampanaCall As Integer
Global vgNUMEROCOMPANIA As String
Global vgNROPOLIZA As String
Global vgNROSECUENCIAL As String
Global vgAPELLIDOYNOMBRE As String
Global vgApellido As String
Global vgNombre As String
Global vgDOMICILIO As String
Global vgDOMICILIONUMERO As String
Global vgLOCALIDAD As String
Global vgPROVINCIA As String
Global vgNUMERODEPOLIZA As String
Global vgCODIGOPOSTAL As String
Global vgNUMEROSECUENCIAL As String
Global vgFECHADESDE As String
Global vgFECHAHASTA As String
Global vgVigenciaVigente As Date
Global vgFECHAVIGENCIA As Date
Global vgFECHAVENCIMIENTO As Date
Global vgFECHAALTAOMNIA As Date
Global vgFECHABAJAOMNIA As Date
Global vgIDAUTO As Integer
Global vgMARCADEVEHICULO As String
Global vgMODELO As String
Global vgCOLOR As String
Global vgAno As String
Global vgPATENTE As String
Global vgTIPODEVEHICULO As Integer
Global vgTipodeServicio As String
Global vgIdTipoDePOliza As Integer
Global vgIDTIPODECOBERTURA As String
Global vgCOBERTURAVEHICULO As String
Global vgCOBERTURAVIAJERO As String
Global vgCOBERTURAHOGAR  As String
Global vgCOBERTURAAP  As String
Global vgTipodeOperacion As String
Global vgOperacion As String
Global vgCATEGORIA As String
Global vgASISTENCIAXENFERMEDAD As String
Global vgCORRIDA As Long
Global vgFECHACORRIDA As Date
Global vgConductor As String
Global vgCodigoDeProductor As String
Global vgCodigoDeServicioVip As String
Global vgTipodeDocumento As String
Global vgNumeroDeDocumento As String
Global vgTipodeHogar As String
Global vgIniciodeAnualidad As String
Global vgPolizaIniciaAnualidad As String
Global vgNroMotor As String
Global vgGama As String
Global vgPosicionRelativa As String
Global vgInformadoSinCobertura As String
Global vgTopeCritales As String
Global vgCodigoDeProceso As String
Global vgFechaDeNacimiento As Date
Global vgOcupacion As String
Global vgImporte As Double
Global vgRama As String
Global vgCertificado As String
Global vgCodigoEnCliente As String
Global vgNroSecunecialEnCliente As String
Global vgorigen As String
Global vgSexo As String
Global vgEmail As String
Global vgEmail2 As String
Global vgPais As String
Global vgTelefono As String
Global vgTelefono2 As String
Global vgTelefono3 As String
Global vgAgencia As String
Global vgIdProducto As String
Global vgDocumentoReferente As String
Global vgDescCampana As String
Global vgReferido As String
Global vgCodigoDeRamo As String
Global vgCuit As String
Global vgMontoCoverturaVidrios As Long
Global vgModificaciones As Long
Global vgCargo As String
Global vgCBU As String
Global vgSucursalCuenta As Long
Global vgFechaModificacion As Date
Global vgCondicionIVA  As Long
Global vgDescripcionDelMundo As String
Global vgPaisOrigen As String
Global vgNroUltimoRecibo As Long
Global vgDescripcionDelModulo As String
Global vgIdLote As Long
Global vgOBSERVACIONES As String
Global vgCambiosDetectados As String
Global vgEmergenciaNombreYApellido As String
Global vgEmergenciaTelefono As String
Global vgEmpresa As String ' se agrego la variable vgEmpresa debido a que esta en la tabla tm_polizas es de tipo "Varchar (50)"
Global vgIdHistorialImportacion As Long
Global vgInformacionAdicionalValor1 As String
Global vgInformacionAdicionalValor2 As String
Global vgInformacionAdicionalValor3 As String
Global vgInformacionAdicionalValor4 As String
Global appPathTemp As String
Global vCoberturaLista(0 To 2, 0 To 20) As String            ' campana,tipoCobertura,coberturaEncontrada
Global vLeidosPorCoberturaLista(0 To 2, 0 To 20) As Long     ' campana,tipoCobertura,coberturaEncontrada
Global vCoberturaActual(0 To 2) As String



'----FIN Variables De Importacion de Texto-----------------------------------------

Public Function LoguearError(ByRef vErr As ErrObject, ByRef vfln As TextStream, ByVal idcampana As Integer, ByVal vHoja As String, ByVal Linea As Integer, ByVal vCampo As String) As Integer
'Usar este codigo en los procesos de lectura
'Dim vCantDeErrores As Integer
'Dim sFileErr As New FileSystemObject
'Dim flnErr As TextStream
'Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
'fln.WriteLine "Errores"
'            vCantDeErrores = 0
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgIdCampana, oSheet.Name, lRow, sName)

LoguearError = 0 'log de error de lectura de las columnas
        If vErr.Number <> 0 Then
            vfln.WriteLine "error " & vErr.Description & " en linea " & Linea & " de la hoja " & vHoja & " columna " & vCampo & " REGISTRO NO IMPORTADO."
            vErr.Clear
            LoguearError = 1
        End If
        
End Function

Public Function LoguearErrorDeConcepto(vComentario As String, ByRef vfln As TextStream, ByVal idcampana As Integer, ByVal vHoja As String, ByVal Linea As Integer, ByVal vCampo As String) As Integer
'Usar este codigo en los procesos de lectura
'Dim vCantDeErrores As Integer
'Dim sFileErr As New FileSystemObject
'Dim flnErr As TextStream
'Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(FileImportacion, 1, Len(FileImportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
'fln.WriteLine "Errores"
'            vCantDeErrores = 0
'                            vCantDeErrores = vCantDeErrores + LoguearError(Err, fln, vgIdCampana, oSheet.Name, lRow, sName)

LoguearErrorDeConcepto = 0 'log de error de lectura de las columnas
vfln.WriteLine "error " & vComentario & " en linea " & Linea & " de la hoja " & vHoja & " coluumna " & vCampo & " REGISTRO NO IMPORTADO."
LoguearErrorDeConcepto = 1
        
End Function

Public Function ValorDistintivo(Valores As String, Scar As Integer) As Long
Dim CtrSumaDe2 As String, ascNum As String
Dim SumaDe1 As Long, SumaDe2 As Long
'ValorDistintivo = 9999999
'Do Until Len(CStr(ValorDistintivo)) <= Scar
    If Len(Valores) Mod 2 <> 0 Then Valores = Valores & "0"
    SumaDe1 = 0
    SumaDe2 = 0
    CtrSumaDe2 = ""
    For i = 1 To Len(Valores)
        ascNum = Mid(Valores, i, 1)
        If IsNumeric(ascNum) Then
            SumaDe1 = SumaDe1 + Asc(ascNum)
            CtrSumaDe2 = CtrSumaDe2 & CStr(Asc(ascNum))
            If i Mod 2 = 0 Then
                SumaDe2 = SumaDe2 + CLng(CtrSumaDe2) - (SumaDe1)
                CtrSumaDe2 = ""
            End If
        End If
    Next
    ValorDistintivo = SumaDe2 - SumaDe1
'    Valores = ValorDistintivo
    Do Until Len(CStr(ValorDistintivo)) < Scar
        sumdigitos = 0
        For j = 1 To Len(CStr(ValorDistintivo)) Step Scar - 1
            sumdigitos = sumdigitos + CInt(Mid(ValorDistintivo, j, Scar - 1))
        Next
        ValorDistintivo = sumdigitos
    Loop
'Loop
End Function

Public Sub Blanquear()
'    vgIDPOLIZA = ""
    vgNUMEROCOMPANIA = ""
    vgNROPOLIZA = ""
    vgNROSECUENCIAL = ""
    vgNROCHASIS = ""
    vgAPELLIDOYNOMBRE = ""
    vgDOMICILIO = ""
    vgLOCALIDAD = ""
    vgPROVINCIA = ""
    vgCODIGOPOSTAL = ""
    vgFECHAVIGENCIA = "00:00:00"
    vgFECHAVENCIMIENTO = "00:00:00"
'    vgFECHAALTAOMNIA = ""
'    vgFECHABAJAOMNIA = ""
    vgIDAUTO = 0
    vgMARCADEVEHICULO = ""
    vgMODELO = ""
    vgCOLOR = ""
    vgAno = 0
    vgPATENTE = ""
    vgTIPODEVEHICULO = 0
    vgTipodeServicio = ""
    vgIDTIPODECOBERTURA = ""
    vgCOBERTURAVEHICULO = ""
    vgCOBERTURAVIAJERO = ""
    vgCOBERTURAHOGAR = ""
    vgTipodeOperacion = ""
    vgOperacion = ""
    vgCATEGORIA = ""
    vgASISTENCIAXENFERMEDAD = ""
'    vgCORRIDA = ""
'    vgFECHACORRIDA = ""
'    vgIdCampana = ""
    vgConductor = ""
    vgCodigoDeProductor = ""
    vgCodigoDeServicioVip = ""
    vgTipodeDocumento = ""
    vgNumeroDeDocumento = ""
    vgTipodeHogar = ""
    vgIniciodeAnualidad = ""
    vgPolizaIniciaAnualidad = ""
    vgTelefono = ""
    vgNroMotor = ""
    vgGama = ""
'    vgPosicionRelativa = ""
'    vgInformadoSinCobertura = ""
    vgTopeCritales = ""
    vgCodigoDeProceso = ""
    vgFechaDeNacimiento = "00:00:00"
    vgOcupacion = ""
    vgImporte = 0
    vgRama = ""
    vgCertificado = ""
    vgCodigoEnCliente = ""
    vgNroSecunecialEnCliente = ""
    vgorigen = ""
    vgCargo = ""
    vgEmpresa = ""
    vgSexo = ""
    vgEmail = ""
    vgEmail2 = ""
    vgPais = ""
    vgTelefono = ""
    vgTelefono2 = ""
    vgTelefono3 = ""
    vgAgencia = ""
    vgIdProducto = ""
    vgDocumentoReferente = ""
    vgReferido = ""
    vgCodigoDeRamo = ""
    vgInformacionAdicionalValor1 = ""
    vgInformacionAdicionalValor2 = ""
    vgInformacionAdicionalValor3 = ""
    vgInformacionAdicionalValor4 = ""
    vCoberturaActual(0) = ""
    vCoberturaActual(1) = ""
    vCoberturaActual(2) = ""

End Sub
Public Function AAMMDDToDD_MM_AA(tfecha As String) As Date

    AAMMDDToDD_MM_AA = DateSerial(Mid(tfecha, 1, 4), Mid(tfecha, 5, 2), Mid(tfecha, 7, 2))

End Function

Public Function mToChar(lValor)
    If lValor < 26 Then
        mToChar = Chr(65 + lValor Mod 26)
    Else
        mToChar = Chr(65 + lValor \ 26 - 1) & Chr(65 + lValor Mod 26)
    End If
End Function

Public Function ComaToPunto(lImporte)
Dim vImpEnt, vImpDec, sImpote, lComaPOs
    If Not IsNumeric(lImporte) Then
        ComaToPunto = CDbl("0.00")
        Exit Function
    End If
    sImporte = CStr(lImporte)
    lComaPOs = InStr(sImporte, ",")
    If lComaPOs > 0 Then
        vImpDec = Mid(sImporte, lComaPOs + 1, Len(sImporte))
        vImpEnt = Mid(sImporte, 1, lComaPOs - 1)
    Else
        vImpDec = "00"
        vImpEnt = sImporte
    End If

    ComaToPunto = (vImpEnt & "." & vImpDec)

End Function


Public Sub RunPackage1(sdir As String, sZone As String)
'Run the package stored in file C:\DTS_UE\TestPkg\VarPubsFields.dts.
'Dim objPackage      As DTS.Package2
'Dim objStep         As DTS.Step
'Dim objTask         As DTS.Task
'Dim objExecPkg      As DTS.ExecutePackageTask
'
'
'    On Error GoTo PackageError
'    sZone = "seteo package"
'    Set objPackage = New DTS.Package
'    Set mobjPkgEvents = objPackage
'    objPackage.FailOnError = True
'
'    'Create the step and task. Specify the package to be run, and link the step to the task.
'    sZone = "seteo Object"
'    Set objStep = objPackage.Steps.New
'    Set objTask = objPackage.Tasks.New("DTSExecutePackageTask")
'    Set objExecPkg = objTask.CustomTask
'    sZone = "seteo EXEC"
'    With objExecPkg
'        .PackagePassword = "user"
'        .FileName = App.Path & "\..\" & sdir '& "\" & "importaPolizasConfiguracion1.dts"
'        .Name = "ExecPkgTask"
'    End With
'    sZone = "seteo Step"
'    With objStep
'        .TaskName = objExecPkg.Name
'        .Name = "ExecPkgStep"
'        .ExecuteInMainThread = True
'    End With
'    sZone = "Agrego Properties"
'    objPackage.Steps.Add objStep
'    objPackage.Tasks.Add objTask
'
'    'Run the package and release references.
'    sZone = "Ejecuta package"
'    Dim pResult As DTSTaskExecResult
'    objPackage.Execute
'
'    sZone = "anula el objeto package"
'    Set objExecPkg = Nothing
'    Set objTask = Nothing
'    Set objStep = Nothing
'    Set mobjPkgEvents = Nothing
'
'    objPackage.UnInitialize
'
'    Exit Sub
'
'PackageError:
'        If Err.Number = 91 Then
'            MsgBox "No cambió el nombre del archivo"
'        Else
'            MsgBox Err.Description
'        End If
'
    
End Sub

'Public Function validoFechas(ByVal fechaNacimiento As DateTime) As Boolean
'    Dim fechaMinima As DateTime
'    Dim hoy As DateTime
'    fechaMinima = "1900/01/01"
'    hoy = Now
'    If (fechaNacimiento < fechaMinima) Then
'        fechaNacimiento = fechaMinima
'    End If
'End Sub


Public Sub listoParaProcesar()
Dim frecuencia As Long
Dim duración As Long
Dim i As Integer
i = 0
frecuencia = 1200 'Hz
duracion = 1000 'ms (1/2 seg.)
Do Until i > 2
Beep frecuencia, duracion
i = i + 1
Loop
End Sub

Public Sub Procesado()
Dim frecuencia As Long
Dim duración As Long
Dim i As Integer
i = 0
frecuencia = 1800 'Hz
duracion = 600 'ms (1/2 seg.)
Do Until i > 2
Beep frecuencia, duracion
i = i + 1
Loop
End Sub

Public Sub StoreProcedureSeteoRegistrosProcesados()

'============Finaliza Proceso, version con vector======================================================se tilda el importador==

            'se declaran el vector que cargara las campañas
            Dim vectorCampana()
            Dim rsVector As New Recordset
            Dim i As Integer

            i = 0


            'se incia un loop para que el file que contiene los registros leidos
            Do Until tf.AtEndOfStream <> True

                ' se hace una query al servidor para obtener los numeros de campaña vigentes
                ssql = " select distinct idcampana from tm_polizas where idcia=" & vgidCia
                rsVector.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

                'a partir del recordeset que guarda el resultado de la consultado se hace un loop para guardar los valores
                'del recordset en el vector.
                Do Until rsVector.EOF
                
                    vectorCampana(i) = rsVector("idcampana").Value
                    i = i + 1
                    rsVector.MoveNext
                Loop

                'se inicia un "for" para recorrer el vector de campañas y ejecutar el store procedure que carga los registros
                'de lectura y registros procesados, levantados luego por el reporte de importaciones.
                
                For i = 0 To UBound(vectorCampana)
                    lidCampana = vectorCampana(i)
                    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lidCampana & ", " & vgCORRIDA
                Next i

            Loop

        

' se agrega el codigo anterior para verificar si de esta manera se corrige el reporte de importacion ( con numero de elctura, fecha de finalizacion de lectura, etc.)
'                            cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxcia " & lIdCia & ", " & vgCORRIDA
'                            Procesado


End Sub

Public Sub TablaTemporal()
Dim ssql As String
On Error Resume Next

cn.Execute "DROP TABLE bandejadeentrada.[dbo].ImportaDatos" & vgidCampana
If Err.Number = -2147217865 Then
    Err.Clear
End If
    ssql = "CREATE TABLE bandejadeentrada.[dbo].[ImportaDatos" & vgidCampana & "]( "
    ssql = ssql & " [IDPOLIZA] [int] NOT NULL, "
    ssql = ssql & " [IDCIA] [Int] NULL , "
    ssql = ssql & " [NUMEROCOMPANIA] [varchar](3) NULL, "
    ssql = ssql & " [NROPOLIZA] [varchar](20) NULL, "
    ssql = ssql & " [NROSECUENCIAL] [varchar](3) NULL, "
    ssql = ssql & " [APELLIDOYNOMBRE] [varchar](255) NULL, "
    ssql = ssql & " [DOMICILIO] [varchar](255) NULL, "
    ssql = ssql & " [LOCALIDAD] [varchar](100) NULL, "
    ssql = ssql & " [PROVINCIA] [varchar](100) NULL, "
    ssql = ssql & " [CODIGOPOSTAL] [varchar](10) NULL, "
    ssql = ssql & " [FECHAVIGENCIA] [datetime] NULL, "
    ssql = ssql & " [FECHAVENCIMIENTO] [datetime] NULL, "
    ssql = ssql & " [FECHAALTAOMNIA] [datetime] NULL, "
    ssql = ssql & " [FECHABAJAOMNIA] [datetime] NULL, "
    ssql = ssql & " [IDAUTO] [int] NULL, "
    ssql = ssql & " [IDPRODUCTO] [int] NULL, "
    ssql = ssql & " [FechadeNacimiento] [datetime] NULL,"
    ssql = ssql & " [MARCADEVEHICULO] [nvarchar](50) NULL, "
    ssql = ssql & " [MODELO] [varchar](50) NULL, "
    ssql = ssql & " [COLOR] [varchar](30) NULL, "
    ssql = ssql & " [ANO] [varchar](4) NULL, "
    ssql = ssql & " [PATENTE] [varchar](15) NULL, "
    ssql = ssql & " [TIPODEVEHICULO] [int] NULL, "
    ssql = ssql & " [TipodeServicio] [char](4) NULL, "
    ssql = ssql & " [IDTIPODECOBERTURA] [int] NULL, "
    ssql = ssql & " [COBERTURAVEHICULO] [char](2) NULL, "
    ssql = ssql & " [COBERTURAVIAJERO] [char](2) NULL, "
    ssql = ssql & " [COBERTURAHOGAR] [char](2) NULL, "
    ssql = ssql & " [COBERTURAAP] [char](2) NULL, "
    ssql = ssql & " [TipodeOperacion] [char](10) NULL, "
    ssql = ssql & " [Operacion] [char](10) NULL, "
    ssql = ssql & " [CATEGORIA] [varchar](2) NULL, "
    ssql = ssql & " [ASISTENCIAXENFERMEDAD] [char](2) NULL, "
    ssql = ssql & " [CORRIDA] [int] NULL, "
    ssql = ssql & " [FECHACORRIDA] [datetime] NULL, "
    ssql = ssql & " [IdCampana] [int] NULL, "
    ssql = ssql & " [PAIS] [varchar] (20) NULL, "
    ssql = ssql & " [Email] [varchar] (70) NULL, "
    ssql = ssql & " [Email2] [varchar] (100) NULL, "
    ssql = ssql & " [Cargo] [varchar] (50) NULL, "
    ssql = ssql & " [Empresa] [varchar] (50) NULL, "
    ssql = ssql & " [Agencia] [nvarchar] (30) NULL, "
    ssql = ssql & " [Conductor] [varchar](50) NULL, "
    ssql = ssql & " [CodigoEnCliente] [varchar](15) NULL, "
    ssql = ssql & " [CodigoDeProductor] [varchar](10) NULL, "
    ssql = ssql & " [CodigoDeServicioVip] [varchar](1) NULL, "
    ssql = ssql & " [TipodeDocumento] [varchar](35) NULL, "
    ssql = ssql & " [NumeroDeDocumento] [varchar](15) NULL, "
    ssql = ssql & " [Documento] [varchar](15) NULL, "
    ssql = ssql & " [TipodeHogar] [char](2) NULL, "
    ssql = ssql & " [IniciodeAnualidad] [char](10) NULL, "
    ssql = ssql & " [PolizaIniciaAnualidad] [char](10) NULL, "
    ssql = ssql & " [Telefono] [char](20) NULL, "
    ssql = ssql & " [NroMotor] [nvarchar](50) NULL, "
    ssql = ssql & " [Sexo] [char](1) NULL, "
    ssql = ssql & " [Gama] [nvarchar](50) NULL, "
    ssql = ssql & " [IdLote] [int] NOT NULL, "
    ssql = ssql & " [InformadoSinCobertura] [nvarchar](1) NULL, "
    ssql = ssql & " [MontoCoverturaVidrios] [float] NULL, "
    ssql = ssql & " [CodigoDeProceso] [nvarchar](1) NULL, "
    ssql = ssql & " [IdTipodePoliza] [char](1) NULL, "
    ssql = ssql & " [EmergenciaNombreYApellido] [varchar] NULL, "
    ssql = ssql & " [EmergenciaTelefono] [varchar] (15) NULL, "
    ssql = ssql & " [Modificaciones] [int] NOT NULL, "
    ssql = ssql & " [ModificacionesEnFechas] [int] NULL,  " ' antes estaba como NOT NULL
    ssql = ssql & " [OBSERVACIONES] [nvarchar](64) NULL, "
    ssql = ssql & " [Cambios Detectados] [int] NULL, "
    ssql = ssql & " [Cuit] [varchar](13) NULL, "
    ssql = ssql & " [InformacionAdicionalValor1] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor2] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor3] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor4] [varchar](50) NULL "
    ssql = ssql & " ) ON [PRIMARY] "
 cn.Execute ssql
End Sub
Public Sub TablaTemporalOptimizadaPrueba()
Dim ssql As String
On Error Resume Next

cn.Execute "DROP TABLE bandejadeentrada.[dbo].ImportaDatos" & vgidCampana
If Err.Number = -2147217865 Then
    Err.Clear
End If
    ssql = "CREATE TABLE bandejadeentrada.[dbo].[ImportaDatos" & vgidCampana & "]( "
    ssql = ssql & " [IDPOLIZA] [int] NOT NULL, "
    ssql = ssql & " [IDCIA] [Int] NULL , "
    ssql = ssql & " [NUMEROCOMPANIA] [varchar](3) NULL, "
    ssql = ssql & " [NROPOLIZA] [varchar](20) not NULL primary key nonclustered, "
    ssql = ssql & " [NROSECUENCIAL] [varchar](3) NULL, "
    ssql = ssql & " [APELLIDOYNOMBRE] [varchar](255) NULL, "
    ssql = ssql & " [DOMICILIO] [varchar](255) NULL, "
    ssql = ssql & " [LOCALIDAD] [varchar](100) NULL, "
    ssql = ssql & " [PROVINCIA] [varchar](100) NULL, "
    ssql = ssql & " [CODIGOPOSTAL] [varchar](10) NULL, "
    ssql = ssql & " [FECHAVIGENCIA] [datetime] NULL, "
    ssql = ssql & " [FECHAVENCIMIENTO] [datetime] NULL, "
    ssql = ssql & " [FECHAALTAOMNIA] [datetime] NULL, "
    ssql = ssql & " [FECHABAJAOMNIA] [datetime] NULL, "
    ssql = ssql & " [IDAUTO] [int] NULL, "
    ssql = ssql & " [IDPRODUCTO] [int] NULL, "
    ssql = ssql & " [FechadeNacimiento] [datetime] NULL,"
    ssql = ssql & " [MARCADEVEHICULO] [nvarchar](50) NULL, "
    ssql = ssql & " [MODELO] [varchar](50) NULL, "
    ssql = ssql & " [COLOR] [varchar](30) NULL, "
    ssql = ssql & " [ANO] [varchar](4) NULL, "
    ssql = ssql & " [PATENTE] [varchar](15) NULL, "
    ssql = ssql & " [TIPODEVEHICULO] [int] NULL, "
    ssql = ssql & " [TipodeServicio] [char](4) NULL, "
    ssql = ssql & " [IDTIPODECOBERTURA] [int] NULL, "
    ssql = ssql & " [COBERTURAVEHICULO] [char](2) NULL, "
    ssql = ssql & " [COBERTURAVIAJERO] [char](2) NULL, "
    ssql = ssql & " [COBERTURAHOGAR] [char](2) NULL, "
    ssql = ssql & " [COBERTURAAP] [char](2) NULL, "
    ssql = ssql & " [TipodeOperacion] [char](10) NULL, "
    ssql = ssql & " [Operacion] [char](10) NULL, "
    ssql = ssql & " [CATEGORIA] [varchar](2) NULL, "
    ssql = ssql & " [ASISTENCIAXENFERMEDAD] [char](2) NULL, "
    ssql = ssql & " [CORRIDA] [int] NULL, "
    ssql = ssql & " [FECHACORRIDA] [datetime] NULL, "
    ssql = ssql & " [IdCampana] [int] NULL, "
    ssql = ssql & " [PAIS] [varchar] (20) NULL, "
    ssql = ssql & " [Email] [varchar] (70) NULL, "
    ssql = ssql & " [Email2] [varchar] (100) NULL, "
    ssql = ssql & " [Cargo] [varchar] (50) NULL, "
    ssql = ssql & " [Empresa] [varchar] (50) NULL, "
    ssql = ssql & " [Agencia] [nvarchar] (30) NULL, "
    ssql = ssql & " [Conductor] [varchar](50) NULL, "
    ssql = ssql & " [CodigoEnCliente] [varchar](15) NULL, "
    ssql = ssql & " [CodigoDeProductor] [varchar](10) NULL, "
    ssql = ssql & " [CodigoDeServicioVip] [varchar](1) NULL, "
    ssql = ssql & " [TipodeDocumento] [varchar](35) NULL, "
    ssql = ssql & " [NumeroDeDocumento] [varchar](15) NULL, "
    ssql = ssql & " [Documento] [varchar](15) NULL, "
    ssql = ssql & " [TipodeHogar] [char](2) NULL, "
    ssql = ssql & " [IniciodeAnualidad] [char](10) NULL, "
    ssql = ssql & " [PolizaIniciaAnualidad] [char](10) NULL, "
    ssql = ssql & " [Telefono] [char](20) NULL, "
    ssql = ssql & " [NroMotor] [nvarchar](50) NULL, "
    ssql = ssql & " [Sexo] [char](1) NULL, "
    ssql = ssql & " [Gama] [nvarchar](50) NULL, "
    ssql = ssql & " [IdLote] [int] NOT NULL, "
    ssql = ssql & " [InformadoSinCobertura] [nvarchar](1) NULL, "
    ssql = ssql & " [MontoCoverturaVidrios] [float] NULL, "
    ssql = ssql & " [CodigoDeProceso] [nvarchar](1) NULL, "
    ssql = ssql & " [IdTipodePoliza] [char](1) NULL, "
    ssql = ssql & " [EmergenciaNombreYApellido] [varchar] NULL, "
    ssql = ssql & " [EmergenciaTelefono] [varchar] (15) NULL, "
    ssql = ssql & " [Modificaciones] [int] NOT NULL, "
    ssql = ssql & " [ModificacionesEnFechas] [int] NULL,  " ' antes estaba como NOT NULL
    ssql = ssql & " [OBSERVACIONES] [nvarchar](64) NULL, "
    ssql = ssql & " [Cambios Detectados] [int] NULL, "
    ssql = ssql & " [Cuit] [varchar](13) NULL, "
    ssql = ssql & " [InformacionAdicionalValor1] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor2] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor3] [varchar](50) NULL, "
    ssql = ssql & " [InformacionAdicionalValor4] [varchar](50) NULL "
    ssql = ssql & " ) WITH (memory_optimized=on, durability=schema_and_data) "
 cn.Execute ssql
End Sub

Public Sub TablaTemporal_UniversalUT()
Dim ssql As String
On Error Resume Next

cn2.Execute "DROP TABLE [UniversalT].[dbo].ImportaDatos" & vgidCampanaCall 'campania en Marge
If Err.Number = -2147217865 Then
    Err.Clear
End If
    ssql = "CREATE TABLE [UniversalT].[dbo].[ImportaDatos" & vgidCampanaCall & "]( "
    ssql = ssql & " [IDCONTACTO] [int] NOT NULL, "
    ssql = ssql & " [IDUTEMPRESA] [Int] NOT NULL , "
    ssql = ssql & " [IDCAMPANA] [Int] NULL , "
    ssql = ssql & " [ApellidoYNombre] [varchar](100) NULL, "
    ssql = ssql & " [FECHAALTA] [datetime] NULL, "
    ssql = ssql & " [CALLE] [varchar](80) NULL, "
    ssql = ssql & " [ALTURA] [varchar](10) NULL, "
    ssql = ssql & " [PISO] [varchar](10) NULL, "
    ssql = ssql & " [DPTO] [varchar](10) NULL, "
    ssql = ssql & " [LOCALIDAD] [varchar](80) NULL, "
    ssql = ssql & " [PROVINCIA] [varchar](80) NULL, "
    ssql = ssql & " [CP] [varchar](8) NULL, "
    ssql = ssql & " [NroDocumento] [varchar](14) NULL, "
    ssql = ssql & " [TELEFONO] [varchar] (60) NULL, "
    ssql = ssql & " [DFECHANAC] [datetime] NULL, "
    ssql = ssql & " [IDPRODUCTO] [int] NULL, "
    ssql = ssql & " [EMAIL] [varchar] (100) NULL, "
    ssql = ssql & " [CORRIDA] [int] NULL, "
    ssql = ssql & " [FECHACORRIDA] [datetime] NULL, "
    ssql = ssql & " [FECHAMOD] [datetime] NULL, "
    ssql = ssql & " [IdLote] [int] NOT NULL, "
    ssql = ssql & " [Sexo] [varchar](1) NULL, "
    ssql = ssql & " ) ON [PRIMARY] "
 cn2.Execute ssql
End Sub

Public Sub CantidadPorCobertura(lidHistorial As Long, sTipoCobertura As String, sCobertura As String, sCantidad As Long, lidProducto As Long)
' recibir 0 en lidProducto si no se utiliza
Dim ssql As String
Dim rs As New Recordset

'Dim cantidadRegistrosLeidos As Long

ssql = "SELECT * FROM TM_ImportacionHistorialCoberturas "
ssql = ssql & " WHERE idHistorialDeImportacion = " & lidHistorial
ssql = ssql & " and TipoCobertura = '" & sTipoCobertura & "' "
ssql = ssql & " and Cobertura = '" & sCobertura & "' "
'ssql = ssql & " and IdProducto = " & lidProducto

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    'cantidadRegistrosLeidos = rs("Leidas") + 1
    ssql = "UPDATE TM_ImportacionHistorialCoberturas "
    ssql = ssql & " SET Leidas = " & sCantidad & " "
    ssql = ssql & " WHERE idHistorialDeImportacion = " & lidHistorial & " "
    ssql = ssql & " and TipoCobertura = '" & sTipoCobertura & "' "
    ssql = ssql & " and Cobertura = '" & sCobertura & "' "
    cn1.Execute ssql
Else
    ssql = "INSERT into TM_ImportacionHistorialCoberturas ("
    ssql = ssql & " idHistorialDeImportacion, "
    ssql = ssql & " TipoCobertura, "
    ssql = ssql & " Cobertura, "
    ssql = ssql & " Leidas) "
    
    ssql = ssql & " values("
    ssql = ssql & lidHistorial & ", '"
    ssql = ssql & sTipoCobertura & "', '"
    ssql = ssql & sCobertura & "', "
    'ssql = ssql & " and IdProducto = " & lidProducto & " "
    ssql = ssql & sCantidad & ")"
    cn1.Execute ssql
End If

rs.Close
    
End Sub

Public Sub CoberturasProcesadas(lCorrida As Long, lidCampana As Integer, lidHistorial As Long)

Dim ssql As String
Dim ssqlHC As String ' ssql para historial de coberturas
Dim rs As New Recordset
Dim cantidadProcesados As Long
Dim vCobertura As String

'======== COBERTURAVEHICULO ========

ssql = " SELECT DISTINCT COBERTURAVEHICULO, "
ssql = ssql & " count(*) as Procesados "
ssql = ssql & " FROM tm_polizas "
'ssql = ssql & " WHERE FECHABAJAOMNIA is null "
ssql = ssql & " WHERE corrida = " & lCorrida
ssql = ssql & " And idcampana = " & lidCampana
ssql = ssql & " group by COBERTURAVEHICULO"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    
    If Not IsNull(rs("COBERTURAVEHICULO")) Then ' podrias reemplazarse por un continue
    
        cantidadProcesados = rs("Procesados")
        vCobertura = rs("COBERTURAVEHICULO")
        
        If vCobertura = "" Then Exit Do
        
        ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
        ssqlHC = ssqlHC & " SET Procesadas = " & cantidadProcesados
        ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
        ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAVEHICULO' "
        ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
        cn1.Execute ssqlHC
    
    End If
    
    rs.MoveNext
    
Loop

rs.Close

'======== COBERTURAVIAJERO ========

ssql = " SELECT DISTINCT COBERTURAVIAJERO, "
ssql = ssql & " count(*) as Procesados "
ssql = ssql & " FROM tm_polizas "
'ssql = ssql & " WHERE FECHABAJAOMNIA is null "
ssql = ssql & " WHERE corrida = " & lCorrida
ssql = ssql & " And idcampana = " & lidCampana
ssql = ssql & " group by COBERTURAVIAJERO"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    
    If Not IsNull(rs("COBERTURAVIAJERO")) Then ' podrias reemplazarse por un continue
    
        cantidadProcesados = rs("Procesados")
        vCobertura = rs("COBERTURAVIAJERO")
        
        If vCobertura = "" Then Exit Do
        
        ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
        ssqlHC = ssqlHC & " SET Procesadas = " & cantidadProcesados
        ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
        ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAVIAJERO' "
        ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
        cn1.Execute ssqlHC
        
    End If
    
    rs.MoveNext
    
Loop

rs.Close

'======== COBERTURAHOGAR ========

ssql = " SELECT DISTINCT COBERTURAHOGAR, "
ssql = ssql & " count(*) as Procesados "
ssql = ssql & " FROM tm_polizas "
'ssql = ssql & " WHERE FECHABAJAOMNIA is null "
ssql = ssql & " WHERE corrida = " & lCorrida
ssql = ssql & " And idcampana = " & lidCampana
ssql = ssql & " group by COBERTURAHOGAR"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF

    If Not IsNull(rs("COBERTURAHOGAR")) Then ' podrias reemplazarse por un continue
        
        cantidadProcesados = rs("Procesados")
        vCobertura = rs("COBERTURAHOGAR")
        
        If vCobertura = "" Then Exit Do
        
        ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
        ssqlHC = ssqlHC & " SET Procesadas = " & cantidadProcesados
        ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
        ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAHOGAR' "
        ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
        cn1.Execute ssqlHC
    
    End If
    
    rs.MoveNext
    
Loop

rs.Close


End Sub

Public Sub CoberturasBajas(lCorrida As Long, lidCampana As Integer, lidHistorial As Long)

'select distinct COBERTURAVEHICULO, count(*) as Bajas
'From tm_polizas
'where idcampana=1077 and corrida <> 2012391 and fechabajaomnia > '2019-08-07 14:18:18.620'
'group by COBERTURAVEHICULO

'select fechaProceso from tm_ImportacionHistorial where corrida = 2012391

Dim ssql As String
Dim ssqlHC As String ' ssql para historial de coberturas
Dim rs As New Recordset
Dim cantidadBajas As Long
Dim vCobertura As String
Dim vFechaProceso As Date

' se obtiene fecha de proceso
ssql = " SELECT fechaProceso FROM tm_ImportacionHistorial WHERE corrida = " & lCorrida
rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
vFechaProceso = rs("fechaProceso")
rs.Close

'======== COBERTURAVEHICULO ========

ssql = " SELECT DISTINCT COBERTURAVEHICULO, "
ssql = ssql & " count(*) as Bajas "
ssql = ssql & " FROM tm_polizas "
ssql = ssql & " WHERE idcampana = " & lidCampana
ssql = ssql & " and corrida <> " & lCorrida
ssql = ssql & " and DATEDIFF(SECOND, '" & vFechaProceso & "', FECHABAJAOMNIA) > 0"
ssql = ssql & " group by COBERTURAVEHICULO"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    
    cantidadBajas = rs("Bajas")
    'If cantidadBajas = Null Then
    '    cantidadBajas = 0
    'End If
    vCobertura = rs("COBERTURAVEHICULO")
    
    ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
    ssqlHC = ssqlHC & " SET Bajas = " & cantidadBajas
    ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
    ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAVEHICULO' "
    ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
    cn1.Execute ssqlHC
    
    rs.MoveNext
    
Loop

rs.Close

'======== COBERTURAVIAJERO ========

ssql = " SELECT DISTINCT COBERTURAVIAJERO, "
ssql = ssql & " count(*) as Bajas "
ssql = ssql & " FROM tm_polizas "
ssql = ssql & " WHERE idcampana = " & lidCampana
ssql = ssql & " and corrida <> " & lCorrida
ssql = ssql & " and DATEDIFF(SECOND, '" & vFechaProceso & "', FECHABAJAOMNIA) > 0"
ssql = ssql & " group by COBERTURAVIAJERO"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    
    cantidadBajas = rs("Bajas")
    vCobertura = rs("COBERTURAVIAJERO")
    
    ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
    ssqlHC = ssqlHC & " SET Bajas = " & cantidadBajas
    ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
    ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAVIAJERO' "
    ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
    cn1.Execute ssqlHC
    
    rs.MoveNext
    
Loop

rs.Close

'======== COBERTURAHOGAR ========

ssql = " SELECT DISTINCT COBERTURAHOGAR, "
ssql = ssql & " count(*) as Bajas "
ssql = ssql & " FROM tm_polizas "
ssql = ssql & " WHERE idcampana = " & lidCampana
ssql = ssql & " and corrida <> " & lCorrida
ssql = ssql & " and DATEDIFF(SECOND, '" & vFechaProceso & "', FECHABAJAOMNIA) > 0"
ssql = ssql & " group by COBERTURAHOGAR"

rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    
    cantidadBajas = rs("Bajas")
    vCobertura = rs("COBERTURAHOGAR")
    
    ssqlHC = "UPDATE TM_ImportacionHistorialCoberturas "
    ssqlHC = ssqlHC & " SET Bajas = " & cantidadBajas
    ssqlHC = ssqlHC & " WHERE idHistorialDeImportacion = " & lidHistorial
    ssqlHC = ssqlHC & " and TipoCobertura = 'COBERTURAHOGAR' "
    ssqlHC = ssqlHC & " and Cobertura = '" & vCobertura & "' "
    cn1.Execute ssqlHC
    
    rs.MoveNext
    
Loop

rs.Close


End Sub

Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Sub ObtenerCoberturas(vIDCampana As Long, vIdProductoEnCliente As String, vCantidadDeErrores As Integer, vflnErr As TextStream, vRow As Long, vName As String)

Dim rsprod As New Recordset
Dim sssql As String

sssql = "Select COBERTURAVEHICULO, COBERTURAVIAJERO, COBERTURAHOGAR, COBERTURAAP, descripcion from TM_PRODUCTOSMultiAsistencias where idcampana = " & vIDCampana & "  and idproductoencliente = '" & vIdProductoEnCliente & "'"
rsprod.Open sssql, cn1, adOpenForwardOnly, adLockReadOnly

If Not rsprod.EOF Then
    vgCOBERTURAVEHICULO = rsprod("coberturavehiculo")
    vCantidadDeErrores = vCantidadDeErrores + LoguearError(Err, vflnErr, vgidCampana, "", vRow, vName)
    vgCOBERTURAVIAJERO = rsprod("coberturaviajero")
    vCantidadDeErrores = vCantidadDeErrores + LoguearError(Err, vflnErr, vgidCampana, "", vRow, vName)
    vgCOBERTURAHOGAR = rsprod("coberturahogar")
    vCantidadDeErrores = vCantidadDeErrores + LoguearError(Err, vflnErr, vgidCampana, "", vRow, vName)
    vgCOBERTURAAP = rsprod("coberturaap")
    vCantidadDeErrores = vCantidadDeErrores + LoguearError(Err, vflnErr, vgidCampana, "", vRow, vName)
    vgIdProducto = vIdProductoEnCliente
    vCantidadDeErrores = vCantidadDeErrores + LoguearError(Err, vflnErr, vgidCampana, "", vRow, vName)
Else
    vCantidadDeErrores = vCantidadDeErrores + LoguearErrorDeConcepto("Producto Inexistente", vflnErr, vgidCampana, "", vRow, vName)
End If

rsprod.Close

End Sub

Public Sub InicializarCoberturaLista()
For i = 0 To 20
    vCoberturaLista(0, i) = "_"
    vCoberturaLista(1, i) = "_"
    vCoberturaLista(2, i) = "_"
Next i
End Sub

Public Sub LeerCoberturas()

vCoberturaActual(0) = vgCOBERTURAVEHICULO
vCoberturaActual(1) = vgCOBERTURAVIAJERO
vCoberturaActual(2) = vgCOBERTURAHOGAR

Dim coberturaPosicion As Integer
Dim posicion As Integer
    
For coberturaPosicion = 0 To 2
    posicion = 0
    Do While posicion < 20 And Len(Trim(vCoberturaActual(coberturaPosicion))) > 1
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

End Sub

Public Sub ProcesarCoberturasLeidas()

Dim posicion As Integer

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

End Sub
