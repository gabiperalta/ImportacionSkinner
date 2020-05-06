Attribute VB_Name = "FuncionesExcel"
Public Function getMaxFilasyColumnas(ExcelFile As Object)

    Dim resultado(1)

    'OBTIENE NUM DE COLUMNAS Y FILAS EN USO
    Dim num_columnas As Integer
    'num_columnas = excelFile.UsedRange.Columns.Count
    num_columnas = ExcelFile.Cells(1, Columns.Count).End(xlToLeft).Column
    resultado(0) = num_columnas
    '----------------------------------------------
    
    'OBTIENE QUE COLUMNA POSEE MAS FILAS CON REGISTROS Y LO ALMACENA
    Dim ultimaFila As Long
    ultimaFila = 1
    Dim LastRow As Long, i As Long
    For i = 1 To num_columnas + 1
        LastRow = ExcelFile.Cells(Rows.Count, i).End(xlUp).Row
        If LastRow > ultimaFila Then
        ultimaFila = LastRow
        End If
    Next i
    resultado(1) = ultimaFila
    '----------------------------------------------------------------
    
    getMaxFilasyColumnas = resultado

End Function

Public Function validarCampos(campos(), fileExcel As Object, numColumnas)

    Dim numCampos
    numCampos = UBound(campos)

    Dim CampoaChequear
    i = 1
    For i = 0 To numCampos
    
        CampoaChequear = campos(i)
        
        Dim j
        
        For j = 1 To numColumnas + 1
        
'        If j = 91 Then
'            a = a
'        End If

            v = Trim(fileExcel.Cells(1, j))
            
            If UCase(CampoaChequear) = UCase(v) Then
            
                campos(i) = True
                
                Exit For
                
            End If
           
        Next j

    Next i
    
    Dim camposFaltantes() As String
    
    ReDim camposFaltantes(1 To 1) As String
    
    Dim proceso As Boolean
    
    Dim tmp As String
    
    proceso = True
    
    For i = 1 To numCampos
    
        If campos(i) <> True Then
        
            proceso = False
            
            tmp = campos(i)
            
            camposFaltantes(UBound(camposFaltantes)) = tmp 'campos(i).Value
            
            ReDim Preserve camposFaltantes(1 To UBound(camposFaltantes) + 1) As String
            
           'ReDim Preserve camposFaltantes(UBound(camposFaltantes) + 1) As Variant
           'camposFaltantes(UBound(camposFaltantes)) = campos(i).Value
           
        End If
    Next
    
    If proceso = True Then
    
        validarCampos = True
        
    ElseIf proceso = False Then
    
        Dim faltantes As String
        
        Dim o As Integer
        
        Dim temp As String
        
        For o = 1 To UBound(camposFaltantes)
        
                temp = camposFaltantes(o)
                
                faltantes = faltantes + temp + " "
                
        Next o
             
        
            validarCampos = faltantes
    End If

End Function

'=====Set corrida que se almacena en el tm_importacionhistorico====================================
Public Function ControlDeCorrida()
Dim rsCorr As New Recordset
                        cn1.Execute "TM_CargaPolizasLogDeSetCorridas " & lIdCampana & ", " & vgCORRIDA ' lIdCampana y vgcorrida se establecen a partir de una consulta hecha en cmdimportar ( evento que selecciona la campaña en el importador de polizas
                        ssql = "Select max(corrida)corrida from tm_ImportacionHistorial where idcampana = " & lIdCampana & " and Registrosleidos is null"
                        rsCorr.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
                        If rsCorr.EOF Then
                            MsgBox "no se determino la corrida, se detendra el proceso"
            
                        Else
                            vgCORRIDA = rsCorr("corrida")
                            
                            
                            
                        End If
                        
                        ControlDeCorrida = vgCORRIDA
                        End Function
                        






                        


