Attribute VB_Name = "AgrosaltaAnt"
Option Explicit

Public Sub ImportarExelAgrosalta()
Dim ssql As String, rsc As New Recordset
Dim lCol, lRow, lCantCol, ll100
Dim v, sName, rsmax

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

    cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos"
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.ImportaDatos", cn, adOpenKeyset, adLockOptimistic
    lRow = 2
    Do While lRow < 30000
        rsc.AddNew

        For lCol = 1 To lCantCol
            v = Worksheets(1).Range(mToChar(lCol - 1) & lRow).Value
            If lCol = 1 And IsEmpty(v) Then Exit Do
            sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                Case "NOMBRE_ASE"
                    rsc("APELLIDOYNOMBRE").Value = v
                Case "MARCA"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "ANO"
                    rsc("ANO").Value = v
                Case "VIGDES"
                    rsc("FECHAVIGENCIA").Value = v
                Case "VIGHAS"
                    rsc("FECHAVENCIMIENTO").Value = v
                Case "LOCALIDAD"
                Case "CP"
                Case "RENUEVA"
'                Case ""
'                   rsc("IDPOLIZA") = v
'                Case ""
'                   rsc("IDCIA") = v
'                Case ""
'                   rsc("NUMEROCOMPANIA") = v
                Case "NU_POLIZA"
                   rsc("NROPOLIZA") = v
'                Case ""
'                   rsc("NROSECUENCIAL") = v
'                Case ""
'                   rsc("APELLIDOYNOMBRE") = v
                Case "DOMICILIO"
                   rsc("DOMICILIO") = v
                Case "LOCALIDAD"
                   rsc("LOCALIDAD") = v
'                Case ""
'                   rsc("PROVINCIA") = v
                Case "POSTAL"
                   rsc("CODIGOPOSTAL") = v
                Case "VIG_DESDE"
                   rsc("FECHAVIGENCIA") = v
                Case "VIG_HASTA"
                   rsc("FECHAVENCIMIENTO") = v
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
'                Case ""
'                   rsc("COLOR") = v
'                Case ""
'                   rsc("ANO") = v
                Case "NU_DOMINIO"
                   rsc("PATENTE") = v
                Case "TIPO_VEHIC"
                   rsc("TIPODEVEHICULO") = v
'                Case ""
'                   rsc("TipodeServicio") = v
'                Case ""
'                   rsc("IDTIPODECOBERTURA") = v
                Case "COD_COBERT"
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
'                Case ""
'                   rsc("CodigoDeProductor") = v
'                Case ""
'                   rsc("CodigoDeServicioVip") = v
'                Case ""
'                   rsc("TipodeDocumento") = v
'                Case ""
'                   rsc("NumeroDeDocumento") = v
'                Case ""
'                   rsc("TipodeHogar") = v
'                Case ""
'                   rsc("IniciodeAnualidad") = v
'                Case ""
'                   rsc("PolizaIniciaAnualidad") = v
'                Case ""
'                   rsc("Telefono") = v
'                Case ""
'                   rsc("NroMotor") = v
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
    cn1.Execute sSPImportacion
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


