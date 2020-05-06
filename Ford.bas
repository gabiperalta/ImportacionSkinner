Attribute VB_Name = "Ford"

Public Sub ImportarFord()

Dim ssql As String, rsc As New Recordset
Dim ll100 As Long
Dim v, sName
Dim regMod As Long
Dim Ll As Long
Dim nroLinea As Long
Dim LongDeLote As Long
Dim lLote As Long
Dim vLote As Integer
Dim vControlDeModificados As Long
Dim vPosicion As Long
Dim vFile As String
Dim fs As New Scripting.FileSystemObject
Dim tf As Scripting.TextStream, sLine As String
Dim vLinea As Long
Dim vlineasTotales As Long
Dim vCampo As String
Dim NombreDividido() As String
Dim PrimeraParte As String

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana
TablaTemporalPK
'TablaTemporal

    '================
    Ll = 1
    ll100 = 0
    '================
    '=======control de lote============
    LongDeLote = 1000
    nroLinea = 1
    vLote = 1
    vlineasTotales = 0
    '================

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


LeoArchivoTXT fileimportacion, vgCORRIDA, Ll, ll100, vLote, LongDeLote ', vlineasTotales
LeoArchivoTXT "QUERY CARDINAL FP.txt", vgCORRIDA, Ll, ll100, vLote, LongDeLote ', vlineasTotales

'================Control de Leidos================================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
    listoParaProcesar
'=================================================================================
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & "copiando linea " & Ll & Chr(13) & " Procesando los datos"
    If MsgBox("Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos====================================
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'=================================================================================
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    Dim rsCMP As New Recordset
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " Procesando los datos"
        DoEvents
        rsCMP.Close
    Next lLote

    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso========================================================
    cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxCia " & lIdCia & ", " & vgCORRIDA
    Procesado
'=====================================================================================
    ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " FinDeProceso"
    ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF


End Sub

    Public Function LeoArchivoTXT(fileimportacion As String, vgCORRIDA As Long, Ll As Long, ll100 As Long, vLote As Integer, LongDeLote As Long)
    Dim vFile As String
    Dim fs As New Scripting.FileSystemObject
    Dim tf As Scripting.TextStream, sLine As String
    Dim NombreDividido() As String
    Dim PrimeraParte As String
    Dim vLinea As Long
    Dim vCampo As String
    Dim regMod As Long

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
    
    vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
    If Not fs.FileExists(vFile) Then Exit Function
    Set tf = fs.OpenTextFile(vFile, ForReading, True)
    'tf.SkipLine Saltea encabezado
        
        Do Until tf.AtEndOfStream
            vLinea = Ll
            sLine = tf.ReadLine
            If Len(Trim(sLine)) < 5 Then Exit Do
            sLine = Replace(sLine, "'", "")
              
                '====maneja los lotes para corte de importacion========
                nroLinea = nroLinea + 1
'                If Ll = 100 Then
'                    MsgBox "manu"
'                End If
                If nroLinea = LongDeLote + 1 Then
                    vLote = vLote + 1
                    vControlDeModificados = 0
                    nroLinea = 1
                End If
                '======================================================
            Blanquear
            vPosicion = 0
        '==================================================================================================
            vCampo = "VIN"
            vPosicion = vPosicion + 1
                vgNroMotor = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
                vgNROPOLIZA = Right(vgNroMotor, 8)
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        '==================================================================================================
'            vCampo = "FechaInicio"
'            vPosicion = vPosicion + 1
'                 vgFECHAVIGENCIA = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
'            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        '==================================================================================================
            vCampo = "PATENTE"
            vPosicion = vPosicion + 1
                vgPATENTE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        '==================================================================================================
            vCampo = "Nombre del Titular"
            vPosicion = vPosicion + 1
                vgAPELLIDOYNOMBRE = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        '==================================================================================================
            vCampo = "MODELO"
            vPosicion = vPosicion + 1
                vgMODELO = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
            sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        '==================================================================================================
            vCampo = "Clase"
            vPosicion = vPosicion + 1
                vgGama = sLine 'Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
                If vgGama = "A" Then 'ex LIVIANO
                    vgidCampana = 1018
                ElseIf vgGama = "C" Then 'ex PESADO
                    vgidCampana = 1019
                Else
                    MsgBox "Error en linea " & Ll & ", campaña inexistente. Revisar base."
                End If
        '==================================================================================================
        NombreDividido() = Split(fileimportacion, ".")
        PrimeraParte = NombreDividido(0)
        If Right(PrimeraParte, 2) = "FP" Then
            vgCOBERTURAVEHICULO = "02"
            vgCOBERTURAVIAJERO = "02"
        ElseIf Right(PrimeraParte, 4) = "TAES" Then
            vgCOBERTURAVEHICULO = "01"
            vgCOBERTURAVIAJERO = "01"
        End If
        '=================================================================================
        vgMARCADEVEHICULO = "FORD"
        If vgidCampana = 1019 Then
            vgTIPODEVEHICULO = 4
        End If
        '=================================================================================
        If vgAPELLIDOYNOMBRE = "" Then
            vgAPELLIDOYNOMBRE = "NO INFORMA"
        End If
        '=================================================================================
        If vgFECHAVIGENCIA = "00:00:00" Then
            vgFECHAVIGENCIA = Now
        End If
        If vgFECHAVENCIMIENTO = "00:00:00" Then
            vgFECHAVENCIMIENTO = DateAdd("m", 6, Now)
        End If
        '=================================================================================
        Dim rscn1 As New Recordset
        ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  idcia = " & Trim(vgidCia) & " and nropoliza = '" & Trim(vgNROPOLIZA) & "'" 'and TipodeOperacion = '" & Trim(vgTipodeOperacion) & "'"
'        ssql = "select IDPOLIZA,"
'        ssql = ssql & "IDCIA,"
'        ssql = ssql & "NROPOLIZA,"
'        ssql = ssql & "APELLIDOYNOMBRE,"
'        ssql = ssql & "FECHAVIGENCIA,"
'        ssql = ssql & "FECHAVENCIMIENTO,"
'        ssql = ssql & "FECHAALTAOMNIA,"
'        ssql = ssql & "MARCADEVEHICULO,"
'        ssql = ssql & "MODELO,"
'        ssql = ssql & "PATENTE,"
'        ssql = ssql & "NroMotor,"
'        ssql = ssql & "TIPODEVEHICULO,"
'        ssql = ssql & "COBERTURAVEHICULO,"
'        ssql = ssql & "COBERTURAVIAJERO,"
'        ssql = ssql & "COBERTURAHOGAR,"
'        ssql = ssql & "IdCampana,"
'        ssql = ssql & "PatenteNumero,"
'        ssql = ssql & "GAMA"
'        ssql = ssql & " from Auxiliout.dbo.tm_Polizas  where  idcia = " & Trim(vgidCia) & " and nropoliza = '" & Trim(vgNROPOLIZA) & "'"
        Dim vdif As Long
        vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
        rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        vgIDPOLIZA = 0
                If Not rscn1.EOF Then
                    vdif = 0  'setea la variale de control de repetido con modificacion en cero
                    If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                    If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                    If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                    If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                    If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                    If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                    If Trim(rscn1("GAMA")) <> Trim(vgGama) Then vdif = vdif + 1
                    If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                    If Trim(rscn1("TIPODEVEHICULO")) <> Trim(vgTIPODEVEHICULO) Then vdif = vdif + 1
                    If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
                    If vgCOBERTURAHOGAR <> "" Then
                        If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                    End If
                    If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                    If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                    If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                    vgIDPOLIZA = rscn1("idpoliza")
                End If
    
            rscn1.Close
    '=================================================================================================================
            ssql = "select * from bandejadeentrada.dbo.ImportaDatos" & lidCampana & " where nropoliza = '" & Trim(vgNROPOLIZA) & "'"
            rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
            
                If rscn1.EOF Then ' Si es verdadero es porque no encontro el registro en la temporal
                    ssql = "Insert into bandejadeentrada.dbo.ImportaDatos" & lidCampana & "("
                    ssql = ssql & "IDPOLIZA, "
                    ssql = ssql & "IDCIA, "
                    ssql = ssql & "IdCampana, "
                    ssql = ssql & "NROPOLIZA, "
                    ssql = ssql & "NroMotor, "
                    ssql = ssql & "APELLIDOYNOMBRE, "
                    ssql = ssql & "FECHAVIGENCIA, "
                    ssql = ssql & "FECHAVENCIMIENTO, "
                    ssql = ssql & "PATENTE, "
                    ssql = ssql & "MODELO, "
                    ssql = ssql & "GAMA, "
                    ssql = ssql & "MARCADEVEHICULO, "
                    ssql = ssql & "TIPODEVEHICULO, "
                    ssql = ssql & "COBERTURAHOGAR, "
                    ssql = ssql & "COBERTURAVEHICULO, "
                    ssql = ssql & "COBERTURAVIAJERO, "
                    ssql = ssql & "CORRIDA, "
                    ssql = ssql & "IdLote, "
                    ssql = ssql & "Modificaciones)"
                    ssql = ssql & " values("
                    ssql = ssql & Trim(vgIDPOLIZA) & ", "
                    ssql = ssql & Trim(vgidCia) & ", "
                    ssql = ssql & Trim(vgidCampana) & ", '"
                    ssql = ssql & Trim(vgNROPOLIZA) & "', '"
                    ssql = ssql & Trim(vgNroMotor) & "', '"
                    ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
                    ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
                    ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
                    ssql = ssql & Trim(vgPATENTE) & "', '"
                    ssql = ssql & Trim(vgMODELO) & "', '"
                    ssql = ssql & Trim(vgGama) & "', '"
                    ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
                    ssql = ssql & Trim(vgTIPODEVEHICULO) & "', '"
                    ssql = ssql & Trim(vgCOBERTURAHOGAR) & "', '"
                    ssql = ssql & Trim(vgCOBERTURAVEHICULO) & "', '"
                    ssql = ssql & Trim(vgCOBERTURAVIAJERO) & "', "
                    ssql = ssql & Trim(vgCORRIDA) & ", '"
                    ssql = ssql & Trim(vLote) & "', '"
                    ssql = ssql & Trim(vdif) & "') "
                    cn.Execute ssql
                    
                Else

                    ssql = "UPDATE bandejadeentrada.dbo.ImportaDatos" & lidCampana & " set IDPOLIZA = " & vgIDPOLIZA & ", "
                    ssql = ssql & " IDCIA = " & vgidCia & ", "
                    ssql = ssql & " IdCampana = " & vgidCampana & ", "
                    ssql = ssql & " NROPOLIZA = '" & vgNROPOLIZA & "', "
                  '  ssql = ssql & " NroMotor = '" & vgNroMotor & "', "
                  '  ssql = ssql & " APELLIDOYNOMBRE = '" & vgAPELLIDOYNOMBRE & "', "
                    ssql = ssql & " FECHAVIGENCIA = '" & vgFECHAVIGENCIA & "', "
                    ssql = ssql & " FECHAVENCIMIENTO = '" & vgFECHAVENCIMIENTO & "', "
                  '  ssql = ssql & " PATENTE = '" & vgPATENTE & "', "
                  '  ssql = ssql & " MODELO = '" & vgMODELO & "', "
                  '  ssql = ssql & " GAMA = '" & vgGama & "', "
                  '  ssql = ssql & " MARCADEVEHICULO = '" & vgMARCADEVEHICULO & "', "
                    ssql = ssql & " TIPODEVEHICULO = '" & vgTIPODEVEHICULO & "', "
                    ssql = ssql & " COBERTURAHOGAR = '" & vgCOBERTURAHOGAR & "', "
                    ssql = ssql & " COBERTURAVEHICULO = '" & vgCOBERTURAVEHICULO & "', "
                    ssql = ssql & " COBERTURAVIAJERO = '" & vgCOBERTURAVIAJERO & "', "
                    ssql = ssql & " CORRIDA = " & vgCORRIDA & ", "
                    ssql = ssql & " IdLote = '" & vLote & "', "
                    ssql = ssql & " Modificaciones = '" & vdif & "'"
                    ssql = ssql & " where NROPOLIZA = '" & vgNROPOLIZA & "'"
                    cn.Execute ssql
                           
                End If
            rscn1.Close
            
    '========Control de errores=========================================================
            If Err Then
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, vCampo)
                Err.Clear
            End If
    '===============================================================================
            If vdif > 0 Then
                regMod = regMod + 1
            End If
            
            Ll = Ll + 1
            ll100 = ll100 + 1
            If ll100 = 100 Then
                ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll & " de la base " & fileimportacion
            ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
    
                    ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lidCampana & "and corrida =" & vgCORRIDA
                 cn1.Execute ssql
                ll100 = 0
            End If
            DoEvents
        Loop
        
    End Function



Public Function TablaTemporalPK()
Dim ssql As String
On Error Resume Next

cn.Execute "DROP TABLE bandejadeentrada.[dbo].ImportaDatos" & vgidCampana

If Err.Number = -2147217865 Then
    Err.Clear
End If

ssql = "CREATE TABLE bandejadeentrada.[dbo].[ImportaDatos" & vgidCampana & "]( "
'ssql = ssql & " [ID] [int] IDENTITY (1,1) PRIMARY KEY, "
ssql = ssql & " [IDPOLIZA] [int] NOT NULL, "
ssql = ssql & " [IDCIA] [int] NULL, "
ssql = ssql & " [NUMEROCOMPANIA] [varchar](3) NULL, "
ssql = ssql & " [NROPOLIZA] [varchar](20) NOT NULL Primary Key, "
'ssql = ssql & " [NROPOLIZA] [varchar](20) NOT NULL, "
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
ssql = ssql & " [FechadeNacimiento] [datetime] NULL, "
ssql = ssql & " [MARCADEVEHICULO] [nvarchar](100) NULL, "
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
ssql = ssql & " [TipodeOperacion] [char](10) NULL, "
ssql = ssql & " [Operacion] [char](10) NULL, "
ssql = ssql & " [CATEGORIA] [varchar](2) NULL, "
ssql = ssql & " [ASISTENCIAXENFERMEDAD] [char](2) NULL, "
ssql = ssql & " [CORRIDA] [int] NULL, "
ssql = ssql & " [FECHACORRIDA] [datetime] NULL, "
ssql = ssql & " [IdCampana] [int] NULL, "
ssql = ssql & " [PAIS] [varchar](20) NULL, "
ssql = ssql & " [Email] [varchar](80) NULL, "
ssql = ssql & " [Email2] [varchar](100) NULL, "
ssql = ssql & " [Cargo] [varchar](50) NULL, "
ssql = ssql & " [Empresa] [varchar](50) NULL, "
ssql = ssql & " [Agencia] [nvarchar](30) NULL, "
ssql = ssql & " [Conductor] [varchar](50) NULL, "
ssql = ssql & " [CodigoEnCliente] [varchar](15) NULL, "
ssql = ssql & " [CodigoDeProductor] [varchar](10) NULL, "
ssql = ssql & " [CodigoDeServicioVip] [varchar](1) NULL, "
ssql = ssql & " [TipodeDocumento] [varchar](20) NULL, "
ssql = ssql & " [NumeroDeDocumento] [varchar](15) NULL, "
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
ssql = ssql & " [EmergenciaNombreYApellido] [varchar](1) NULL, "
ssql = ssql & " [EmergenciaTelefono] [varchar](15) NULL, "
ssql = ssql & " [Modificaciones] [int] NOT NULL, "
ssql = ssql & " [OBSERVACIONES] [nvarchar](64) NULL, "
ssql = ssql & " [Cambios Detectados] [int] NULL "
ssql = ssql & " ) ON [PRIMARY] "
cn.Execute ssql

End Function
