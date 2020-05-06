Attribute VB_Name = "Morrone"
Option Explicit

Public Sub ImportarMorrone()

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

On Error GoTo errores
    
 
Dim col As New Scripting.Dictionary
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oExcel = New Excel.Application
oExcel.Visible = False
Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
Set oSheet = oBook.Worksheets(1)
    
Dim filas As Integer
Dim columnas As Integer
Dim extremos(1)
columnas = FuncionesExcel.getMaxFilasyColumnas(oSheet)(0)
extremos(1) = FuncionesExcel.getMaxFilasyColumnas(oSheet)(1)

'columnas = extremos(0)
filas = extremos(1)

Dim camposParaValidar(15)
camposParaValidar(0) = "Patente"
camposParaValidar(1) = "Poliza"
camposParaValidar(2) = "marcadevehiculo"
camposParaValidar(3) = "modelo"
camposParaValidar(4) = "color"
camposParaValidar(5) = "Vigencia"
camposParaValidar(6) = "coberturavehiculo"
camposParaValidar(7) = "coberturaviajero"
camposParaValidar(8) = "coberturahogar"
camposParaValidar(9) = "domicilio"
camposParaValidar(10) = "localidad"
camposParaValidar(11) = "provincia"
camposParaValidar(12) = "documento"
camposParaValidar(13) = "nombre"
camposParaValidar(14) = "correlativo"
camposParaValidar(15) = "baja"

If FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas) = True Then

    cn.Execute "DELETE FROM bandejadeentrada.dbo.importaDatosV2 where idcampana=" & 855 & " and idcia=" & 10001297
    rsc.Open "SELECT * FROM bandejadeentrada.dbo.importaDatosV2 where idcampana=" & 855 & " and idcia=" & 1000129, cn, adOpenKeyset, adLockOptimistic
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

    rsc("idcia") = 10001297
    rsc("idcampana") = 855

        For lCol = 1 To columnas
            sName = col.Item(lCol)
            v = oSheet.Cells(lRow, lCol)
        If IsEmpty(v) = False Then
        
'        If UCase(Trim(sName)) <> "VENCIMIENTO" Then
'        MsgBox "1"
'        End If
        
            If lCol = 1 And IsEmpty(v) Then
                Exit Do
            End If
            
            
            
            'sName = col.Item(lCol)
            Select Case UCase(Trim(sName))
                Case "PATENTE"
                    rsc("PATENTE").Value = v
                Case "MARCADEVEHICULO"
                    rsc("MARCADEVEHICULO").Value = v
                Case "MODELO"
                    rsc("MODELO").Value = v
                Case "COLOR"
                   rsc("COLOR") = v
                Case "AÑO"
                    rsc("ANO").Value = v
                Case "POLIZA"
                   rsc("NROPOLIZA") = v
                Case "NOMBRE"
                   rsc("APELLIDOYNOMBRE") = v
                Case "DOMICILIO"
                   rsc("DOMICILIO") = v
                Case "LOCALIDAD"
                   rsc("LOCALIDAD") = v
                Case "PROVINCIA"
                   rsc("PROVINCIA") = v
                Case "DOCUMENTO"
                   rsc("NUMERODEDOCUMENTO") = v
                Case "VIGENCIA"
                    If InStr(1, v, "/") > 0 Then
                        rsc("FECHAVIGENCIA") = v
                    Else
                        rsc("FECHAVIGENCIA") = Mid(v, 7, 2) & "/" & Mid(v, 5, 2) & "/" & Mid(v, 1, 4)
                    End If
                Case "VENCIMIENTO"
                     If InStr(1, v, "/") > 0 Then
                        rsc("FECHAVENCIMIENTO") = v
                    Else
                        
                        'rsc("FECHAVENCIMIENTO") = Mid(v, 7, 2) & "/" & Mid(v, 5, 2) & "/" & Mid(v, 1, 4)
                    End If
                Case "BAJA"
                    rsc("FECHABAJAOMNIA") = v
                 
                Case "COBERTURAHOGAR"
                   rsc("COBERTURAHOGAR") = v
                Case "COBERTURAVIAJERO"
                   rsc("COBERTURAVIAJERO") = v
                Case "COBERTURAVEHICULO"
                   rsc("COBERTURAVEHICULO") = v
                Case "CORRELATIVO"
                   rsc("NROSECUENCIAL").Value = v


            End Select
            
            
            End If
            
           
        
            
            
            
        Next
        
         If IsEmpty(rsc("FECHAVENCIMIENTO")) = True Then
           rsc("FECHAVENCIMIENTO") = DateAdd("yyyy", 1, rsc("fechaVIGENCIA"))
        End If
        
        If IsEmpty(rsc("NROSECUENCIAL")) Or v = "" Then
            rsc("NROSECUENCIAL").Value = 0
        End If
        
        rsc.Update
        lRow = lRow + 1
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2
            ll100 = 0
        End If
        DoEvents

    Loop
    
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & lRow - 2 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ssql = "select max(CORRIDA) as maxCorrida from Auxiliout.dbo.tm_polizas"
    rsUltCorrida.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    vUltimaCorrida = rsUltCorrida("maxCorrida") + 1
    
    cn1.Execute "update tm_campana set UltimaCorridaError='' , UltimaCorridaCantidadderegistros=0  where idcampana=855"
         ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
   
    ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & vTipoDePoliza & Chr(13) & " procesando linea 1" & Chr(13) & " de " & vlineasTotales & " Procesando los datos"
    DoEvents
    cn1.CommandTimeout = 300
    cn1.Execute sSPImportacion
    vIDCampana = 855
    ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidaddeRegistros from tm_campana where idcampana=" & vIDCampana
    rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    vRegistrosProcesados = vRegistrosProcesados + rsCMP("UltimaCorridaCantidaddeRegistros")
    If Trim(rsCMP("UltimaCorridaError")) <> "OK" Then
        MsgBox " msg de Error de proceso : " & rsCMP("UltimaCorridaError")
    End If
    rsCMP.Close
    
    cn1.Execute "update tm_campana set  UltimaCorridaCantidadderegistros = " & vRegistrosProcesados & " where idcampana=" & vIDCampana
    cn.Execute "DELETE FROM bandejadeentrada.dbo.importaDatosV2 where idcampana=" & 855 & " and idcia=" & 10001297
Else
    MsgBox ("Los siguientes campos obligatorios no fueron encontrados: " & FuncionesExcel.validarCampos(camposParaValidar(), oSheet, columnas)), vbCritical, "Error"
End If

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




