Attribute VB_Name = "OrganizacionLaSpina"

Public Sub ImportarOrganizacionLaSpina()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vCtrolPatente As Boolean, vCtrolVigencia As Boolean, vCtrolVencimiento As Boolean
Dim col As New Scripting.Dictionary


cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos507"

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lIdCampana

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

'========inicia Excel y abre el workbook===============================
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion, False, True)
Set oSheet = oBook.Sheets(1)

'========control de lectura del archivo de datos=======================
If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If

'=======inicio del control de corrida==================================
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

'=====================================================================
   
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

'If lCol = 1 Then
'    MsgBox "Faltan campos"
'    Exit Sub
'End If


rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatos507", cn, adOpenKeyset, adLockOptimistic
lRow = 2
Do While lRow < 30000

    For lCol = 1 To lCantCol
       'v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
        v = oSheet.Cells(lRow, lCol)
        If lCol = 1 And IsEmpty(v) Then Exit Do
        sName = col.Item(lCol)
        Select Case UCase(Trim(sName))
            Case "PATENTE"
                vgPATENTE = v
                vgNROPOLIZA = v
                vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "", lRow, sName)
            Case "NOMBRE"
                vgAPELLIDOYNOMBRE = v
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

        End Select
    Next
    
    '==============  IMPORTANTE   =======================================================
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
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & lIdCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "' "
       rscn1.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
       vdif = 1  'setea la variale de control en 1 por si es un registro que no existe si existe luego pone modificacion en cero
       vgIDPOLIZA = 0
               If Not rscn1.EOF Then
                   vdif = 0  'setea la variale de control de repetido con modificacion en cero
                   If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                   If Trim(rscn1("NROPOLIZA")) <> Trim(vgNROPOLIZA) Then vdif = vdif + 1
                   If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                   If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                   If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                   If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                   If Trim(rscn1("FECHAVIGENCIA")) <> Trim(vgFECHAVIGENCIA) Then vdif = vdif + 1
                   If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                   vgIDPOLIZA = rscn1("idpoliza")
    '                If vdif > 0 Then 'bloque para identificar modificaciones al hacer un debug.
    '                End If
                   
               End If
    
        rscn1.Close
            
    '========insert que se hace a la tabla temporal que se crea al comienzo======================
            
    ssql = "Insert into bandejadeentrada.dbo.ImportaDatos507 ("
    ssql = ssql & "IDPOLIZA, "
    ssql = ssql & "PATENTE, "
    ssql = ssql & "NROPOLIZA, "
    ssql = ssql & "APELLIDOYNOMBRE, "
    ssql = ssql & "MARCADEVEHICULO, "
    ssql = ssql & "MODELO, "
    ssql = ssql & "ANO, "
    ssql = ssql & "FECHAVIGENCIA, "
    ssql = ssql & "FECHAVENCIMIENTO, "
    ssql = ssql & "Modificaciones)"
    ssql = ssql & " values("
    ssql = ssql & Trim(vgIDPOLIZA) & ", '"
    ssql = ssql & Trim(vgPATENTE) & "', '"
    ssql = ssql & Trim(vgNROPOLIZA) & "', '"
    ssql = ssql & Trim(vgAPELLIDOYNOMBRE) & "', '"
    ssql = ssql & Trim(vgMARCADEVEHICULO) & "', '"
    ssql = ssql & Trim(vgMODELO) & "', '"
    ssql = ssql & Trim(vgAno) & "', '"
    ssql = ssql & Trim(vgFECHAVIGENCIA) & "', '"
    ssql = ssql & Trim(vgFECHAVENCIMIENTO) & "', '"
    ssql = ssql & Trim(vdif) & "') "
    cn.Execute ssql
    
            
    '========Control de errores===========================================
    If Err Then
        vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", lRow, "")
        Err.Clear
    End If
    
    If vdif > 0 Then
        regMod = regMod + 1
    End If
        
    lRow = lRow + 1
    ll100 = ll100 + 1
    If ll100 = 100 Then
        ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow

    '========update ssql para % de modificaciones segun leidos en reporte de importaciones=========================================================
        ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (lRow) & ",  parcialModificaciones =" & regMod & " where idcampana=" & lIdCampana & "and corrida =" & vgCORRIDA
                cn1.Execute ssql
        ll100 = 0
    End If
    DoEvents
Loop

'================Control de Leidos===========llama al storeprocedure para hacer un update en tm_importacionHistorial
cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & lRow
listoParaProcesar


ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

'===============inicio del Control de Procesos===========================================
cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
DoEvents

cn1.Execute sSPImportacion & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso========================================================
cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
Procesado
'=====================================================================================
ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & lRow & Chr(13) & " FinDeProceso"
ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF



oExcel.Workbooks.Close
Set oExcel = Nothing

End Sub

