Attribute VB_Name = "ResguardoAsistencias"
Option Explicit

Public Sub ImportarExelResguardoAsistencias()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax
Dim vUltimaCorrida As Long
Dim rsUltCorrida As New Recordset
Dim vIDCampana As Long
Dim vidTipoDePoliza As Long
Dim vTipoDePoliza As String
Dim vRegistrosProcesados As Long
Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim sArchivo As String
Dim rsCMP As New Recordset


Dim col As New Scripting.Dictionary
'Dim mExcel As New Excel.Application
'Dim wb
        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet

On Error GoTo errores
        ' Inicia Excel y abre el workbook
        Set oExcel = New Excel.Application
        oExcel.Visible = False
        Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
        Set oSheet = oBook.Worksheets(1)
'Dim sh As Excel.Sheets
    'Set mExcel = CreateObject("Excel.Application")
'    oExcel.Visible = False
'    Set oBooks = oExcel.Workbooks.Open(App.Path & "\" & sDirImportacion & "\" & FileImportacion, False, True)
    vIDCampana = 524
        v = " "
    lCol = 1
    lRow = 1
    Do While lCol < 50
        v = oSheet.Range(mToChar(lCol - 1) & "1").Value
        If IsEmpty(v) Then Exit Do
        sName = v
        col.Add lCol, v
        lCol = lCol + 1
    Loop
    lCantCol = lCol

    If lCol = 1 Then
        MsgBox "Faltan campos"
        Exit Sub
    End If

    cn.Execute "DELETE FROM bandejadeentrada.dbo.importaDatos"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.importaDatos", cn, adOpenKeyset, adLockOptimistic
    lRow = 3
    Do While lRow < 30000
        rsc.AddNew

        For lCol = 1 To lCantCol
            v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            If lCol = 1 And IsEmpty(v) Then
                Exit Do
            End If
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                Case "Marca y Modelo"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "AÑO"
                    rsc("ANO").Value = v
                Case "RENUEVA"
'                Case ""
'                   rsc("IDPOLIZA") = v
'                Case ""
'                   rsc("IDCIA") = v
'                Case ""
'                   rsc("NUMEROCOMPANIA") = v
                Case "Nº Cliente"
                   rsc("NROPOLIZA") = v
'                Case ""
'                   rsc("NROSECUENCIAL") = v
                Case "Apellido"
                   rsc("APELLIDOYNOMBRE") = v
                Case "NOMBRE"
                   rsc("APELLIDOYNOMBRE") = rsc("APELLIDOYNOMBRE") & ", " & v
                Case "Dirección"
                   rsc("DOMICILIO") = v
                Case "LOCALIDAD"
                   rsc("LOCALIDAD") = v
                Case "PROVINCIA"
                    Select Case v
                        Case "B"
                            v = "Buenos Aires"
                        Case "C"
                            v = "Capital Federal"
                    End Select
                   rsc("PROVINCIA") = v
                Case "Cód.Postal"
                   rsc("CODIGOPOSTAL") = v
                Case "F.Inicio"
                    If InStr(1, v, "/") > 0 Then
                        rsc("FECHAVIGENCIA") = v
                    Else
                        rsc("FECHAVIGENCIA") = Mid(v, 7, 2) & "/" & Mid(v, 5, 2) & "/" & Mid(v, 1, 4)
                    End If
                Case "F.Final"
                     If InStr(1, v, "/") > 0 Then
                        rsc("FECHAVENCIMIENTO") = v
                    Else
                        rsc("FECHAVENCIMIENTO") = Mid(v, 7, 2) & "/" & Mid(v, 5, 2) & "/" & Mid(v, 1, 4)
                    End If
'                Case ""
'                   rsc("FECHAALTAOMNIA") = v
'                Case ""
'                   rsc("FECHABAJAOMNIA") = v
'                Case ""
'                   rsc("IDAUTO") = v
'                Case ""
'                   rsc("MARCADEVEHICULO") = v
'                Case ""
'                   rsc("MODELO") = v
                Case "Color"
                   rsc("COLOR") = v
'                Case ""
'                   rsc("ANO") = v
                Case "TIPO_VEHIC"
                    If v = "A" Then
                        v = "01"
                    ElseIf v = "P" Then
                        v = "03"
                    Else
                        v = "00"
                    End If

                   rsc("TIPODEVEHICULO") = v
'                Case ""
'                   rsc("TipodeServicio") = v
'                Case ""
'                   rsc("IDTIPODECOBERTURA") = v
                Case "SCOBERTURA"
                    If v = "A" Then
                        v = "01"
                    ElseIf v = "B" Then
                        v = "02"
                    ElseIf v = "B1" Then
                        v = "02"
                    ElseIf v = "C" Then
                        v = "02"
                    ElseIf v = "C1" Then
                        v = "02"
                    Else
                        v = "00"
                    End If
                   rsc("COBERTURAVEHICULO") = v
'                Case ""
'                   rsc("COBERTURAVIAJERO") = v
'                Case ""
'                   rsc("TipodeOperacion") = v
'                Case ""
'                   rsc("Operacion") = v
'                Case ""
'                   rsc("CATEGORIA") = v
'                Case ""
'                   rsc("ASISTENCIAXENFERMEDAD") = v
'                Case ""
'                   rsc("CORRIDA") = v
'                Case ""
'                   rsc("FECHACORRIDA") = v
'                Case ""
'                   rsc("IdCampana") = v
'                Case ""
'                   rsc("Conductor") = v
                Case "CodProductor"
                   rsc("CodigoDeProductor") = v
'                Case ""
'                   rsc("CodigoDeServicioVip") = v
                Case "Tipo Doc."
                    Select Case v
                        Case "C"
                            v = "2"
                        Case "D"
                            v = "1"
                        Case "T"
                            v = "5"
                    End Select
                   rsc("TipodeDocumento") = v
                Case "Nº Documento"
                   rsc("NumeroDeDocumento") = v
'                Case ""
'                   rsc("TipodeHogar") = v
'                Case ""
'                   rsc("IniciodeAnualidad") = v
'                Case ""
'                   rsc("PolizaIniciaAnualidad") = v
                Case "Teléfono"
                   rsc("Telefono") = v
                Case "Motor"
                   rsc("NroMotor") = v
'                Case ""
'                   rsc("Gama") = v
'                Case ""
'                   rsc("NroDocumento") = v
            End Select
        Next
        rsc.Update
        lRow = lRow + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtProcesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2
            ll100 = 0
        End If
        DoEvents

    Loop
    oExcel.Workbooks.Close
    ImportadordePolizas.txtProcesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    
    cn1.Execute "update tm_campana set UltimaCorridaError='' , UltimaCorridaCantidadderegistros=0  where idcampana=" & vIDCampana
    
    ImportadordePolizas.txtProcesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    DoEvents
    cn1.CommandTimeout = 300
    cn1.Execute sSPImportacion
    ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros from tm_campana where idcampana=" & vIDCampana
    rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    vRegistrosProcesados = vRegistrosProcesados + rsCMP("UltimaCorridaCantidaddeRegistros")
    If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
        MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
    End If
    rsCMP.Close
    
    cn1.Execute "update tm_campana set  UltimaCorridaCantidadderegistros = " & vRegistrosProcesados & " where idcampana=" & vIDCampana
Exit Sub
errores:
    oExcel.Workbooks.Close
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow & " Columna: " & sName
    End If



End Sub


