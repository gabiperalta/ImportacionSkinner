Attribute VB_Name = "TriunfoAuxilio25"
Option Explicit
Public Sub ImportarAuxilio25()
Dim ssql As String, rsc As New Recordset, rs As New Recordset
Dim lCol, lRow, lCantCol, ll100
'Dim v
Dim sName, rsmax
Dim vCtrolIdentificacion As Boolean, vCtrolVigencia As Boolean
Dim vCtrolVencimiento As Boolean, vCtroltipodeasistencia As Boolean, vCtrolidproductoasistencia As Boolean
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
'Dim nroLinea As Long
Dim vCampo As String
Dim vPosicion As Long
'Dim lLote As Long
'Dim vLote As Long
Dim rsUltCorrida As New Recordset
Dim vUltimaCorrida As Long
'Dim rsCMP As New Recordset
'Dim LongDeLote As Integer
Dim vlineasTotales As Long
Dim vLTIPODEVEHICULO As String
Dim vTipodeServicio As String
Dim vTipodeServicioActual As String
Dim vCS As String
Dim vtipodeasistencia As String
Dim vCobertura As String
Dim vFile As String
Dim vArchivo As String
Dim vNroPolizaB As String
Dim regMod As Long


'==========Aqui setear el Caracter de separacion=========
    vCS = Chr(9) '";"
'=======================================================

On Error Resume Next

Dim vOBSERVACIONES As String
vhorainicio = Now
'Dim mExcel As New Excel.Application
'Dim wb
'        Dim oExcel As Excel.Application
'        Dim oBook As Excel.Workbook
'        Dim oSheet As Excel.Worksheet

'On Error GoTo errores

'    Dim oExcel As Object
'    Dim oBook As Object
'    Dim oSheet As Object
'
'    Set oExcel = New Excel.Application
'    Set oBook = oExcel.Workbooks.Open(App.Path & vgPosicionRelativa & sDirImportacion & "\" & FileImportacion, False, True)
'
'    oBook.Activate
'    vArchivo = Mid(FileImportacion, 1, InStr(1, FileImportacion, ".")) & "csv"
'    ActiveWorkbook.SaveAs FileName:= _
'        App.Path & vgPosicionRelativa & sDirImportacion & "\" & vArchivo, FileFormat _
'        :=xlCSV, CreateBackup:=False

cn.Execute "DELETE FROM bandejadeentrada.dbo.ImportaDatos510"

On Error Resume Next
vgidCia = lIdCia
vgidCampana = lidCampana

Dim vCantDeErrores As Integer
Dim sFileErr As New FileSystemObject
Dim flnErr As TextStream
Set flnErr = sFileErr.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & Mid(fileimportacion, 1, Len(fileimportacion) - 5) & "_" & Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & Minute(Now) & Second(Now) & ".log", True)
flnErr.WriteLine "Errores"
vCantDeErrores = 0

If Err Then
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End If


Dim sFile As New FileSystemObject
Dim fln As TextStream
Set fln = sFile.CreateTextFile(App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion & Year(Now) & Month(Now) & Day(Now) & ".log", True)
fln.WriteLine "Linea; Certificado; Documento; Patente; Fecha; Detalle"
    
'======Genera el control de productos================================================================
    vProductosPosibles = ""
    ssql = "Select Distinct categoria from tm_Servicioxcia where idcia = " & vgidCia
    rs.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
    vProductosPosibles = rs("Categoria")
    rs.MoveNext
    Do Until rs.EOF
        vProductosPosibles = vProductosPosibles & "," & rs("Categoria")
        rs.MoveNext
    Loop
'============================se crea  un file de lectura "tf"==========================================
  
        
    Ll = 0
    vFile = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
    If Not fs.FileExists(vFile) Then Exit Sub
    Set tf = fs.OpenTextFile(vFile, ForReading, True)
    sLine = tf.ReadLine 'si no trae fila de titulos eliminar este read
 '======='control de lectura del archivo de datos=======================
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Exit Sub
    End If
'=====inicio del control de corrida====================================
    Dim rsCorr As New Recordset
    cn1.Execute "TM_CargaPolizasLogDeSetCorridasxcia " & lIdCia & ", " & vgCORRIDA ' se usa este storeprocedure debido a que la compañia tiene varias campañas.
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

    
'=========================================== Comienzo del Do de lectura ======================================================
   Do Until tf.AtEndOfStream
        Ll = Ll + 1
        vLinea = Ll
        sLine = tf.ReadLine
        If Len(Trim(sLine)) < 5 Then Exit Do
        sLine = Replace(sLine, "'", "*")
    
'=======Control de Lote===============================
        nroLinea = nroLinea + 1
        If nroLinea = LongDeLote + 1 Then
            vLote = vLote + 1
            nroLinea = 1
        End If
'=====================================================
        vPosicion = 0
        
        vCampo = "tipodeasistencia"
        vPosicion = vPosicion + 1
        vtipodeasistencia = Mid(sLine, 1, InStr(1, sLine, vCS) - 1)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
   
            If vtipodeasistencia = "MOTOS" Then
                vgTIPODEVEHICULO = 5
                vgidCampana = 510
                vgidCia = vgidCia
            ElseIf vtipodeasistencia = "AUTOS" Then
                vgTIPODEVEHICULO = 1
                vgidCampana = 734
                vgidCia = vgidCia
            ElseIf vtipodeasistencia = "CAMIONES" Then
                vgTIPODEVEHICULO = 4
                vgidCampana = 931
                vgidCia = vgidCia
            ElseIf vtipodeasistencia = "BICICLETAS" Then
                vgTIPODEVEHICULO = 8
                vgidCampana = 980
                vgidCia = vgidCia
            
            End If

'==================================================================================================
        vCampo = "idasistenciacertificado"
        vPosicion = vPosicion + 1
            vCertificado = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
            vgNROPOLIZA = vCertificado
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "localidad"
        vPosicion = vPosicion + 1
            vgLOCALIDAD = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "provincia"
        vPosicion = vPosicion + 1
            vgPROVINCIA = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "fechaemision"
        vPosicion = vPosicion + 1
        vgFECHAVIGENCIA = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If Not IsDate(vgFECHAVIGENCIA) Then
                    vgFECHAVIGENCIA = ""
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & vCampo
                End If
'==================================================================================================
        vCampo = "email"
        vPosicion = vPosicion + 1
        vgEmail = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
'==================================================================================================
        vCampo = "documento"
        vPosicion = vPosicion + 1
        vDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
            vgNumeroDeDocumento = vDocumento
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
'==================================================================================================
        vCampo = "tipodedocumento"
        vPosicion = vPosicion + 1
            vgTipodeDocumento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
    
'==================================================================================================
        vCampo = "fechadenacimiento"
        vPosicion = vPosicion + 1
            If IsDate(Mid(sLine, 1, InStr(1, sLine, vCS) - 1)) Then
                vgFechaDeNacimiento = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
            Else
                vError = 1
                fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & " ; " & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & vCampo
            End If

        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If Not IsDate(vgFechaDeNacimiento) Then
                    vgFechaDeNacimiento = ""
               ' Else
                 '   vError = 1
                '    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & vCampo
                End If
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "vigenciafin"
        vPosicion = vPosicion + 1
            vgFECHAVENCIMIENTO = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If Not IsDate(vgFECHAVENCIMIENTO) Then
                    vgFECHAVENCIMIENTO = ""
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & vCampo
                End If
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "observaciones"
        vPosicion = vPosicion + 1
            vOBSERVACIONES = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
            vOBSERVACIONES = Mid(vgDOMICILIO, 1, 64)
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "calle"
        vPosicion = vPosicion + 1
            vgDOMICILIO = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "altura"
        vPosicion = vPosicion + 1
            vgDOMICILIO = vgDOMICILIO & " " & Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "piso"
        vPosicion = vPosicion + 1
            vgDOMICILIO = vgDOMICILIO & " " & Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "barrio"
        vPosicion = vPosicion + 1
            vgDOMICILIO = vgDOMICILIO & " " & Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "partido"
        vPosicion = vPosicion + 1
            vgDOMICILIO = vgDOMICILIO & " " & Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                vgDOMICILIO = Mid(vgDOMICILIO, 1, 80)
'==================================================================================================
        vCampo = "cp4"
        vPosicion = vPosicion + 1
            vgCODIGOPOSTAL = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If Not IsNumeric(vgCODIGOPOSTAL) Or Len(vgCODIGOPOSTAL) > 5 Then
                    vgCODIGOPOSTAL = ""
                End If
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "telcelular"
        vPosicion = vPosicion + 1
        If Len(Mid(sLine, 1, InStr(1, sLine, vCS) - 1)) < 20 Then
            vgTelefono = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        Else
            vgTelefono = Mid(Mid(sLine, 1, InStr(1, sLine, vCS) - 1), 1, 20)
        End If
        
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "anio"
        vPosicion = vPosicion + 1
            vgAno = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If Not IsNumeric(vgAno) Or vgAno < 1900 Or vgAno > Year(Now()) Or Len(vgAno) > 4 Then
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & sName
                    vgAno = 0
                End If
'==================================================================================================
        vCampo = "patente"
        vPosicion = vPosicion + 1
            vgPATENTE = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                'If Len(v) > 10 Then
                '    vgPATENTE = ""
                '    vError = 1
                '    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & sName
                'End If
'==================================================================================================
         vCampo = "nromotor"
        vPosicion = vPosicion + 1
            vgNroMotor = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
         vCampo = "color"
        vPosicion = vPosicion + 1
            vgCOLOR = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "empresaapellidoynombre"
        vPosicion = vPosicion + 1
            vgAPELLIDOYNOMBRE = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "marca"
        vPosicion = vPosicion + 1
            vgMARCADEVEHICULO = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "modelo"
        vPosicion = vPosicion + 1
        vgMODELO = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        vgMODELO = Mid(vgMODELO, 1, 20)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "idproductoasistencia"
        vPosicion = vPosicion + 1
            vCobertura = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
                If InStr(vProductosPosibles, Trim(vCobertura)) > 0 And IsEmpty(vCobertura) = False Then
                    vCobertura = Mid("0" & vCobertura, Len("0" & vCobertura) - 1, 2)
                    vgCOBERTURAVEHICULO = vCobertura
                    vgCOBERTURAVIAJERO = vCobertura
                    vgCOBERTURAHOGAR = vCobertura
                Else
                    vError = 1
                    fln.WriteLine " " & vLinea & " ; " & vCertificado & " ; " & vDocumento & " ; " & vPatente & " ; " & Now & " ; " & "ERROR EN LINEA " & vLinea & " VALOR " & Mid(sLine, 1, InStr(1, sLine, vCS) - 1) & " EN LA COLUMNA " & vCampo
                End If
            
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
         vCampo = "pais"
        vPosicion = vPosicion + 1
            vgPais = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)

'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        vCampo = "vip"
        vPosicion = vPosicion + 1
            vgCodigoDeServicioVip = Trim(Mid(sLine, 1, InStr(1, sLine, vCS) - 1))
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
'==================================================================================================
        'Salta campo sin uso, en caso de necesitar el campo saltado, reemplazar por variable
        vCampo = "Salta Campo"
        vPosicion = vPosicion + 1
        sLine = Mid(sLine, InStr(1, sLine, vCS) + 1)
        
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

 Dim vcamp As Integer
 Dim vdif As Long
' cn1.Execute " TM_NormalizaPolizas " & vgNROPOLIZA & ", " & vNroPolizaB & " OUT"
    ssql = "select *  from Auxiliout.dbo.tm_Polizas  where  IdCampana = " & vgidCampana & " and nroPoliza = '" & Trim(vgNROPOLIZA) & "'"
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
'                If Trim(rscn1("IdTipoDoc")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("TipodeDocumento")) <> Trim(vgTipodeDocumento) Then vdif = vdif + 1
                If Trim(rscn1("FechadeNacimiento")) <> Trim(vgFechaDeNacimiento) Then vdif = vdif + 1
                If Trim(rscn1("FECHAVENCIMIENTO")) <> Trim(vgFECHAVENCIMIENTO) Then vdif = vdif + 1
                If IsDate(rscn1("FECHABAJAOMNIA")) Then vdif = vdif + 1
'                If Trim(rscn1("OBSERVACIONES")) <> Trim(vOBSERVACIONES) Then vdif = vdif + 1
                'If Trim(rscn1("DOMICILIO")) <> Trim(vgDOMICILIO) Then vdif = vdif + 1
                If Trim(rscn1("CODIGOPOSTAL")) <> Trim(vgCODIGOPOSTAL) Then vdif = vdif + 1
                'If Trim(rscn1("Telefono")) <> Trim(vgTelefono) Then vdif = vdif + 1
                If Trim(rscn1("ANO")) <> Trim(vgAno) Then vdif = vdif + 1
                If Trim(rscn1("PATENTE")) <> Trim(vgPATENTE) Then vdif = vdif + 1
                If Trim(rscn1("NroMotor")) <> Trim(vgNroMotor) Then vdif = vdif + 1
                If Trim(rscn1("COLOR")) <> Trim(vgCOLOR) Then vdif = vdif + 1
'                If Trim(rscn1("APELLIDOYNOMBRE")) <> Trim(vgAPELLIDOYNOMBRE) Then vdif = vdif + 1
                If Trim(rscn1("MARCADEVEHICULO")) <> Trim(vgMARCADEVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("MODELO")) <> Trim(vgMODELO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVEHICULO")) <> Trim(vgCOBERTURAVEHICULO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAVIAJERO")) <> Trim(vgCOBERTURAVIAJERO) Then vdif = vdif + 1
                If Trim(rscn1("COBERTURAHOGAR")) <> Trim(vgCOBERTURAHOGAR) Then vdif = vdif + 1
                If Trim(rscn1("PAIS")) <> Trim(vgPais) Then vdif = vdif + 1
                If Trim(rscn1("CodigoDeServicioVip")) <> Trim(vgCodigoDeServicioVip) Then vdif = vdif + 1
                vgIDPOLIZA = rscn1("idpoliza")
'                If vdif > 0 Then
'                vdif = vdif
'                 End If
                
            End If

        rscn1.Close
'-=================================================================================================================
        ssql = "Insert into bandejadeentrada.dbo.ImportaDatos510 ("
        ssql = ssql & "IdPoliza, "
        ssql = ssql & "TIPODEVEHICULO, "
        ssql = ssql & "IdCampana, "
        ssql = ssql & "idcia, "
        ssql = ssql & "NROPOLIZA, "
        ssql = ssql & "LOCALIDAD, "
        ssql = ssql & "PROVINCIA, "
        ssql = ssql & "FECHAVIGENCIA, "
        ssql = ssql & "Email, "
        ssql = ssql & "NumeroDeDocumento, "
        'ssql = ssql & "IdTipoDoc, "
        ssql = ssql & "TipodeDocumento, "
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
        ssql = ssql & "CodigoDeServicioVip, "
        ssql = ssql & "IdLote, "
        ssql = ssql & "CORRIDA, "
        ssql = ssql & "Modificaciones)"
        ssql = ssql & " values("
        ssql = ssql & Trim(vgIDPOLIZA) & ", "
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
        ssql = ssql & Trim(vgCodigoDeServicioVip) & "', '"
        ssql = ssql & Trim(vLote) & "', "
        ssql = ssql & Trim(vgCORRIDA) & ", '"
        ssql = ssql & Trim(vdif) & "') "
        cn.Execute ssql
        
'========Control de errores=========================================================
                If Err Then
                    vCantDeErrores = vCantDeErrores + LoguearError(Err, flnErr, vgidCampana, "Proceso", Ll, "")
                    Err.Clear
                
                End If
'===========================================================================================
        
        ll100 = ll100 + 1
        If ll100 = 100 Then
            ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll
            ''========update ssql para porcentaje de modificaciones segun leidos en reporte de importaciones=========================================================
            If vdif > 0 Then
                regMod = regMod + 1
            End If
            ssql = "update Auxiliout.dbo.tm_ImportacionHistorial set parcialLeidos=" & (Ll) & ",  parcialModificaciones =" & regMod & " where idcia=" & vgidCia & "and corrida =" & vgCORRIDA
            cn1.Execute ssql
            '=========================================================================================================================================================
            ll100 = 0
        End If
        DoEvents
    Loop
    
'================Control de Leidos===============================================
    cn1.Execute "TM_CargaPolizasLogDeSetLeidos " & vgCORRIDA & ", " & Ll
    listoParaProcesar
'=================================================================================
    
    ImportadordePolizas.txtprocesando.Text = "Importando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " copiando linea " & Ll - 1 & Chr(13) & " Procesando los datos"
    If MsgBox("¿Desea Procesar los datos de " & vgDescCampana & " ?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'===============inicio del Control de Procesos===========================================
    cn1.Execute "TM_CargaPolizasLogDeSetInicioDeProceso " & vgCORRIDA
'==================================================================================
    ImportadordePolizas.txtprocesando.BackColor = &HC0C0FF
    
    Dim rsCMP As New Recordset
    DoEvents
    For lLote = 1 To vLote
        cn1.CommandTimeout = 300
        cn1.Execute sSPImportacion & " " & lLote & ", " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana
        ssql = "Select UltimaCorridaError,UltimaCorridaUltimaPoliza from tm_campana where idcampana=" & vgidCampana
        rsCMP.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
        ImportadordePolizas.txtprocesando.Text = "Procesando " & ImportadordePolizas.cmbCia.Text & Chr(13) & " procesando linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " Procesando los datos"
        DoEvents
        rsCMP.Close
    Next lLote

    cn1.Execute "TM_BajaDePolizasControlado" & " " & vgCORRIDA & ", " & vgidCia & ", " & vgidCampana

'============Finaliza Proceso, version con vector======================================================se tilda el importador==

    'se llama a la subrutina que setea los registros procesados ( fecha corrida, etc...) que toma el reporte de importaciones.
    cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxCia " & vgidCia & ", " & vgCORRIDA ' se cambia lIdCampana por vgidcia
    'se llama a la subrutina que tira el veep de finalizacion de registros procesados.
    Procesado
'==================================================================================================================================================================
    
            
''            se declaran el vector que cargara las campañas
'            Dim vectorCampana()
'            Dim rsVector As New Recordset
'            Dim i As Integer
'
'            i = 0
'
'
'            'se incia un loop para que el file que contiene los registros leidos
'            Do Until tf.AtEndOfStream <> True
'
'                ' se hace una query al servidor para obtener los numeros de campaña vigentes
'                ssql = " select distinct idcampana from tm_polizas where idcia=" & vgidCia
'                rsVector.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'
'                'a partir del recordeset que guarda el resultado de la consultado se hace un loop para guardar los valores
'                'del recordset en el vector.
'                Do Until rsVector.EOF
'                    vectorCampana(i) = rsVector("idcampana").Value
'                    i = i + 1
'                    rsVector.MoveNext
'                Loop
'
'                'se inicia un "for" para recorrer el vector de campañas y ejecutar el store procedure que carga los registros
'                'de lectura y registros procesados, levantados luego por el reporte de importaciones.
'                For i = 0 To UBound(vectorCampana)
'                    lIdCampana = vectorCampana(i)
'                    cn1.Execute "TM_CargaPolizasLogDeSetProcesados " & lIdCampana & ", " & vgCORRIDA
'                Next i
'
'            Loop
'
'            'se llama al proceso que tira el veep de finalizacion de registros procesados.
'            Procesado
'
'' se agrega el codigo anterior para verificar si de esta manera se corrige el reporte de importacion ( con numero de elctura, fecha de finalizacion de lectura, etc.)
'                            cn1.Execute "TM_CargaPolizasLogDeSetProcesadosxcia " & lIdCia & ", " & vgCORRIDA
'                            Procesado
'=====================================================================================
        ImportadordePolizas.txtprocesando.Text = "Procesado " & ImportadordePolizas.cmbCia.Text & Chr(13) & " proceso linea " & (lLote * LongDeLote) & Chr(13) & " de " & Ll & " FinDeProceso"
        ImportadordePolizas.txtprocesando.BackColor = &HFFFFFF

Exit Sub

errores:
'    oExcel.Workbooks.Close
'    Set oExcel = Nothing
    vgErrores = 1
    If lRow = 0 Then
        MsgBox Err.Description
    Else
        MsgBox Err.Description & " en linea " & lRow & " Columna: " & sName
    End If

End Sub




