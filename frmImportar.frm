VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ImportadordePolizas 
   BackColor       =   &H00FFFF80&
   Caption         =   "Importador de Polizas"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtprocesando 
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmImportar.frx":0000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   960
   End
   Begin VB.TextBox TxtCoberturas 
      Height          =   1095
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCantidadDepolizas 
      Caption         =   "Cantidad de Polizas"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporte Final"
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   6015
      Begin VB.TextBox TxtActivos 
         Height          =   285
         Left            =   4800
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtRegistros 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCorrida 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Registros Activos"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Registros Procesados"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Corrida"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCampana 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   225
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1515
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmImportar.frx":0083
      Left            =   1440
      List            =   "frmImportar.frx":008D
      TabIndex        =   1
      Top             =   60
      Width           =   3045
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar"
      Height          =   225
      Left            =   4560
      MaskColor       =   &H008080FF&
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label labelS 
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   ":"
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.Label labelM 
      Height          =   615
      Left            =   6960
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   ":"
      Height          =   495
      Left            =   6720
      TabIndex        =   17
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label labelH 
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.Label labelTiempo 
      Caption         =   "Tiempo"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Campaña"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Compania"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "ImportadordePolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'Private WithEvents mobjPkgEvents As DTS.Package
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim h As Integer, m As Integer, s As Integer




Private Sub cmbCia_Click()
Dim rsc As New Recordset, i As Integer
Dim ssql As String

    cmbCampana.Clear
    
    If Trim(cmbCia.Text) = "" Then Exit Sub
    
    ssql = "Select a.idcontacto,b.idcampana,b.DESCCAMPANA campana,b.DirImportacion "
    ssql = ssql & " from tm_Contactos a inner join Tm_Campana b on a.IdContacto = b.idcontacto "
    ssql = ssql & " where a.NombreYApellido = '" & cmbCia.Text & "' and a.idtipodecontacto = 1 and b.activa = 1 order by a.NombreYApellido"
    rsc.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    i = 1
    Do Until rsc.EOF
        cmbCampana.AddItem rsc("campana")
        rsc.MoveNext
    Loop
    

End Sub



Private Sub cmdCantidadDepolizas_Click()
    Dim rsc As New Recordset
    Dim rs As New Recordset
    Dim ssql As String
    
    ssql = "SELECT     p.COBERTURAVEHICULO as Cobertura, COUNT(p.COBERTURAVEHICULO) AS Cantidad"
    ssql = ssql & " From TM_POLIZAS p inner join Tm_Contactos c on c.idcontacto=p.idcia "
    ssql = ssql & " Where c.NombreYApellido = '" & cmbCia.Text & "' And (p.FECHABAJAOMNIA Is Null) And p.COBERTURAVEHICULO > 0"
    ssql = ssql & " GROUP BY p.COBERTURAVEHICULO"
    ssql = ssql & " Union"
    ssql = ssql & " SELECT     p.COBERTURAHogar as Cobertura, COUNT(p.COBERTURAHogar) AS Cantidad"
    ssql = ssql & " From TM_POLIZAS p inner join Tm_Contactos c on c.idcontacto=p.idcia "
    ssql = ssql & " Where c.NombreYApellido = '" & cmbCia.Text & "'  And (p.FECHABAJAOMNIA Is Null) And p.COBERTURAHogar > 0"
    ssql = ssql & " GROUP BY p.COBERTURAHogar"
    rsc.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    
    TxtCoberturas.Text = "Cob    Cantidad" & Chr(13) & Chr(10)
    Do Until rsc.EOF
        TxtCoberturas.Text = TxtCoberturas.Text & rsc("cobertura") & "    " & rsc("Cantidad") & Chr(13) & Chr(10)
        rsc.MoveNext
    Loop
    TxtCoberturas.Visible = True

End Sub

Private Sub cmdImportar_Click()
'Dim sDirImportacion As String, sdtsImportacion As String, sFileImportacion As String
Dim cmm As New Command, d As Date
Dim rsc As New Recordset
Dim rsCorrida As New Recordset
Dim rsCant As New Recordset
Dim rsCamp As New Recordset
'Dim sfile As FileSystemObject
Dim proceso As Boolean, lCantidadRegistros As Long, lCantidadActivos As Long
Dim sZone As String, ssql As String
'Dim lIdCia As Long, lIdCampana As Long
'Dim FileImportacion As String, sSPImportacion As String
Dim sAno As String, sMes As String, sDia As String
Dim sFile, sArchivo As String, sExtencion As String, lImportados As Long
    
Dim fs As New Scripting.FileSystemObject
    
   On Error Resume Next
   
    proceso = False
    ' http://support.microsoft.com/kb/319058/es
    'Dim prm1 As New ADODB.Parameter
    'Dim prm2 As New ADODB.Parameter
    'On Error GoTo errores
    'Me.Width = Me.Width * 2 comente aca tambien
    Blanquear
   'txtProcesando.Width = txtProcesando.Width * 2 esta comenteeeeee

    txtprocesando.Text = "Importando " & cmbCia.Text
    txtprocesando.Visible = True
    DoEvents
    
    ssql = "Select a.idcontacto,a.NombreYApellido,b.idcampana,b.DirImportacion,b.dtsImportacion,b.FileImportacion, "
    ssql = ssql & " b.SPimportacion "
    ssql = ssql & " from tm_Contactos a inner join Tm_Campana b on a.IdContacto = b.idcontacto "
    ssql = ssql & " where a.NombreYApellido = '" & cmbCia.Text & "' and b.DescCampana = '" & cmbCampana.Text & "'"
    rsc.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    If rsc.EOF Then
        txtprocesando.Text = "Epa tio, hay problemas!!!"
        Exit Sub
    End If
    
    lIdCia = rsc("idcontacto")
    lidCampana = rsc("idcampana")
    sDirImportacion = rsc("DirImportacion")
    sdtsImportacion = "" & rsc("dtsImportacion")
    sSPImportacion = "" & rsc("SPimportacion") & "Controlado"
    fileimportacion = rsc("FileImportacion")
    vgDescCampana = cmbCampana.Text
    rsc.Close
    vgErrores = 0
    txtCorrida.Text = ""
    TxtActivos.Text = ""
    txtRegistros.Text = ""
         
'    ssql = "Select c.idcampana, e.idutempresa from tm_campana c "
'    ssql = ssql & " inner join tm_utcontratos uc on c.idcontrato = uc.idcontrato "
'    ssql = ssql & " inner join tm_utEmpresas e on uc.idutempresa= e.idutempresa "
'    ssql = ssql & " where c.desccampana = '" & cmbCampana.Text & "'"
'    rsc.Open ssql, cn2, adOpenForwardOnly, adLockReadOnly
'    If Not rsc.EOF Then
'        lIdCampanaCall = rsc("Idcampana")
'        lIdUtempresaCall = rsc("idutempresa")
'    End If
'    rsc.Close
'    vgErrores = 0
'
    
'If Not creaTabla(vgDescCampana) Then
'    MsgBox "No Creo Nada, fijate"
'End If
'    ssql = "DELETE FROM importaDatosHorizonte"
'    cn.Execute ssql
    'Blanquea
    'ssql = "update Tm_campanas set UltimaCorrida=0, UltimaCorridaCantidadDeRegistros=0 where idcampana = " & lIdCampana
    'cn1.Execute ssql
          
    ImportadordePolizas.Width = 4800 ' oculta los botones de Importar y Salir

    Timer1.Enabled = True
        
    appPathTemp = sDirImportacion
    If Not EndsWith(App.Path, "\") Then
        sDirImportacion = "\" & sDirImportacion
    End If
    
    Select Case appPathTemp
        
        Case "Ace"
            ImportarChubbAce
        Case "Agrosalta"
            ImportarExelAgrosalta
        Case "Morrone"
            ImportarMorrone
        Case "GM_GRUPO_ASEGURADOR"
            ImportarGM
        Case "AgrosaltaMoron"
            ImportarExelAgrosaltaMoron
        Case "El Comercio"
            ImportarElComercio
        Case "PSA_Peugeot"
            ImportarPSAPeugeot
        Case "PSA_Citroen"
            ImportarPSACitroen
        Case "Federal"
            ImportarFederal
        Case "ANTARTIDA SEGUROS"
            ImportarAntartida
        Case "Willassist"
            ImportarWillAssist
        Case "Erika"
            ImportarExelErika
        Case "organizacion Torres"
            ImportarOrganizacionTorres
        Case "Mundi"
            ImportarExelMundi
        Case "Basilio"
            ImportarBasilio
        Case "Boston -TElefonica"
            ImportarBoston
        Case "La Veloz"
            ImportarLaVeloz
        Case "LIDERAR"
            ImportarLiderar
        Case "Liderar_Vargas"
            ImportarLiderarVargas
        Case "Iveco"
            ImportarIveco
        Case "NUEVOBANCOCOMERCIALDEURUGUAY"
            ImportarNUEVOBANCOCOMERCIALDEURUGUAY
        Case "Triunfo Seguros"
            ImportarTriunfoSeguros
        Case "INTERNATIONALASSISTANCE"
            ImportarInternationalAssistance
        Case "IntegritySeguros"
            ImportarIntegritySeguros
        Case "SiempreSegurosGenerales"
            ImportarSiempreSegurosGenerales
        Case "OrganizacionLaSpina"
            ImportarOrganizacionLaSpina
        Case "AGS"
            ImportarAGS
        Case "HorizonteSeguros"
            ImportarHorizonte
        Case "MAYASISTENCIAS"
            ImportarMayAsistencias
        Case "RESGUARDOASISTENCIAS"
            ImportarResguardoAsistencias
        Case "BuenosAiresAsistencias"
            ImportarExelBuenosAiresAsistencias
        Case "CaminosProtegidos"
            ImportarCaminosProtegidos
        Case "TriunfoSegurosAP"
            ImportarTriunfoSegurosAp
        Case "LifeAssistance"
            ImportarExcelLifeAssistance
        Case "TarjetaPlata"
            ImportadorTarjetaPlata
        Case "CRYASOCIADOS"
            ImportarCRyASOCIADOS
        Case "Qualia"
            ImportarQualiaCSV
        Case "AUXILIO24"
            ImportarAuxilio25
        Case "SPCONSULTORA"
            ImportarSPConsultora
        Case "GALENOLIFE"
            ImportarGalenoLife
        Case "Claugel"
            ImportarClaugel
        Case "ASISTENCIAVIP"
            ImportarAsistenciaVip
        Case "UPSA"
            ImportarGenericoExcelSinFechas
        Case "amtmf"
            ImportarGenericoExcel
        Case "AMMMA"
            ImportarGenericoExcel
        Case "AMEP"
            ImportarGenericoExcel
        Case "AMCSudEste"
            ImportarGenericoExcelSinFechas
        Case "ASMEPRIV"
            ImportarGenericoExcel
        Case "AL PRODUCTOR ASESOR AYELEN LARDAPIDE"
            ImportarAlProductorAsesor
        Case "AgenciaDeViajes"
            ImportarGenericoExcel
        Case "ARCOR"
            ImportarGenericoExcel
        Case "HINO"
            ImportarHino
        Case "Huarpes"
            ImportarGenericoExcel
        Case "AMSeSa"
            ImportarGenericoExcel
        Case "AMMetropolitana"
            ImportarGenericoExcel
        Case "CooperativaCardinal"
            ImportarGenericoExcel
        Case "ColonMotos"
            ImportarColonMotos
        Case "ColonEmpleados"
            ImportarGenericoExcel
        Case "NACIONALVIDASEGUROSDEPERSONAS"
            ImportarNacionalVida
        Case "Premedic"
            ImportarPremedic
        Case "Desafio Ruta 40"
            ImportarGenericoExcel
        Case "DTC"
            ImportarExelCHUBB
        Case "BIND"
            ImportarExelCHUBB
        Case "Global Asistencias"
            ImportarGenericoExcel
        Case "BeneficioSeguros"
            ImportarBeneficioSeguros
        Case "RedDeServicios"
            ImportarGenericoExcel
        Case "AMLitoral"
            ImportarGenericoExcel
        Case "ColonAsistencias"
            ImportarColonAsistencias
        Case "ColonSeguros"
            ImportarColonSegurosUnificado
        Case "OBRASOCIALDELAMATANZA"
            ImportarClaugel
        Case "SI SERVICIOS(kodalle)"
            ImportarGenericoExcel
        Case "FORD"
            ImportarFord
        Case "FORTALEZA SEGUROS"
            ImportarFortalezaSeguros
        Case "ArgusSalud"
            ImportarGenericoExcel
        Case "TarjetaActual"
            ImportarGenericoExcel
        Case "CocinaSaludable"
            ImportarGenericoExcel
        Case "AMBULANCIAS - BANCOINDUSTRIAL"
            ImportarBancoIndustrial_Ambulancias
        Case "PROTECCION GLOBAL ( LA ANONIMA )"
            ImportarProteccionGlobalTarjetaAnonima
        Case "BEEFIX"
            ImportarGenericoExcel
        Case "BIND - Area protegida"
            ImportarGenericoExcel
        Case "SEGURCOOP"
            ImportarSegurcoop
        Case "RRHH EMERGENCIAS CARDINAL"
            ImportarGenericoExcel
        Case "Hogar - BancoIndustrial"
            ImportarBancoIndustrial_Hogar
        Case "MEDICASSIST"
            'ImportarBancoIndustrial_Hogar
            ImportarGenericoExcel
        Case "GarantiaDirecta"
            ImportarGenericoExcel
        Case "Assisto"
            ImportarGenericoExcel
        Case "AMSE"
            ImportarGenericoExcel
        Case "Caruso"
            ImportarGenericoExcel
        Case "PACK BENEFICIO SALUD"
            ImportarGenericoExcel
        Case Else
            MsgBox "ERROR #420 - La campana no esta cargada"
        End Select
    
    If Err Then
     MsgBox "Error " & Err.Description
     Err.Clear
'     Exit Sub
    End If

'    MsgBox "error 6", vbInformation, "Error"
'    If vgErrores <> 0 Then
'        MsgBox "Error en los datos leidos"
'        Me.Width = 6240
'        txtProcesando.Width = 6015
'        txtProcesando.Visible = False
'        Exit Sub
'    End If
    
'    If sSPImportacion <> "" Then
'        Dim cm As Command, rs As Recordset
'
'        rsc.Open "SELECT COUNT(*) AS TOTAL FROM BANDEJADEENTRADA.DBO.importaDatosAGS", cn, adOpenKeyset, adLockReadOnly
''        rsc.Open "SELECT COUNT(*) AS TOTAL FROM IMPORTADATOS", cn, adOpenKeyset, adLockReadOnly
'        lImportados = rsc!Total.Value
'        rsc.Close
'
'        Enabled = True
'    End If
    
'    ssql = "Select UltimaCorrida,UltimaCorridaError,UltimaCorridaUltimaPoliza,UltimaCorridaCantidadDeRegistros,UltimaCorridaCantidadImportada from tm_campana where idcampana=" & lIdCampana
'    rsCamp.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'    If Not rsCamp.EOF Then
'            txtCorrida.Text = rsCamp("UltimaCorrida")
'            If Not IsNull(rsCamp("UltimaCorridaCantidadImportada")) Then txtRegistros.Text = rsCamp("UltimaCorridaCantidadImportada")
'            TxtActivos.Text = rsCamp("UltimaCorridaCantidadDeRegistros")
'    End If
'
'    Dim vCorrida As Long
'    If Not IsNull(rsCamp("ultimacorrida")) Then
'    vCorrida = rsCamp("UltimaCorrida")
'    End If
'
'    If Not rsCamp.EOF Then
'        If Not IsNull(rsCamp("UltimaCorridaError")) Then
'
'            ssql = "Select count(*) as CantidadRegistros from tm_Polizas "
'            ssql = ssql & " WHERE CORRIDA in (Select max(corrida) from tm_Polizas WHERE IDCAMPANA = " & lIdCampana & " and datediff(min,fechacorrida,getdate())< 5) "
'            ssql = ssql & " AND IDCAMPANA=" & lIdCampana
'            rsCorrida.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'
'                If rsCorrida.State <> 0 Then
'                    lCantidadRegistros = rsCorrida("CantidadRegistros")
'                End If
''
'            If lCantidadRegistros = 0 Then
'                vCorrida = 0
'            Else
'                'sCorrida = rsCamp("UltimaCorrida")
'                'sCorrida = vCorrida
'            End If
'
'            txtProcesando.Text = "Resultado del Proceso :" & rsCamp("UltimaCorridaError") & " Ultima poliza procesada: " & rsCamp("UltimaCorridaUltimaPoliza") & " En la corrida Nro: " & vCorrida & " Cantidad de Polizas Importadas: " & lCantidadRegistros & " De " & rsCamp("UltimaCorridaCantidadDeRegistros") & " De registros procesados"
'        End If
'    End If
'    proceso = True
'
    sAno = Year(Date)
    sMes = Month(Date) 'if(len(Month(Date))=1;"0" & cstr(Month(Date));Month(Date))
    sDia = Day(Date)
    
'    If proceso Then
        sZone = "Scripting:"
        Set sFile = CreateObject("Scripting.FileSystemObject")

'            If rsCorrida.State <> 0 Then rsCorrida.Close
'            ssql = "Select count(*) as CantidadRegistros from tm_Polizas WHERE CORRIDA in (Select max(corrida) from tm_Polizas WHERE IDCAMPANA = " & lIdCampana & ")"  'AND IDCAMPANA=" & lIdCampana
'            rsCorrida.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'
'            If rsCorrida.State <> 0 Then
'                 lCantidadRegistros = rsCorrida("CantidadRegistros")
'
'            End If
'
'            If rsCorrida.State <> 0 Then rsCorrida.Close
'            ssql = "Select count(*) as CantidadActivos "
'            ssql = ssql & " from tm_Polizas WHERE  IDCAMPANA=" & lIdCampana & " and fechabajaomnia is null "
'            rsCorrida.Open ssql, cn1, adOpenForwardOnly, adLockReadOnly
'
'            If rsCorrida.State <> 0 Then
'                 lCantidadActivos = rsCorrida("CantidadActivos")
'
'            End If
'

Dim pathArchivo As String
Dim pathArchivoNuevo As String
Dim nroArchivo As Integer
nroArchivo = 1

        'RENOMBRA EL ARCHIVO AL FINALIZAR SEGURCOOP
        If vgidCampana = 1073 Then
            sExtencion = ".txt"
            For nroArchivo = 0 To 8
                
                Select Case nroArchivo
                    Case 0
                        sArchivo = "ATM"
                    Case 1
                        sArchivo = "viaj"
                    Case 2
                        sArchivo = "COMBINADO"
                    Case 3
                        sArchivo = "INTEGRAL"
                    Case 4
                        sArchivo = "AUTO_A"
                    Case 5
                        sArchivo = "AUTO_B"
                    Case 6
                        sArchivo = "AUTO_C"
                    Case 7
                        sArchivo = "AUTO_E"
                    Case 8
                        sArchivo = "CAMIONES"
                    Case Else
                        Exit For
                End Select
                
                pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & sExtencion
                pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
                sFile.CopyFile pathArchivo, pathArchivoNuevo
                sZone = "Delete:"
                sFile.DeleteFile pathArchivo
                txtCorrida.Text = vgCORRIDA
                If Not IsNull(rsCamp("UltimaCorridaCantidadImportada")) Then txtRegistros.Text = rsCamp("UltimaCorridaCantidadImportada")
                txtRegistros.Text = lCantidadRegistros
                TxtActivos.Text = lCantidadActivos
                
            Next
            
            
'            fileimportacion = "ATM.txt"
'            sArchivo = Mid(fileimportacion, 1, Len(fileimportacion) - 4)
'            sExtencion = Mid(fileimportacion, InStr(1, fileimportacion, "."), Len(fileimportacion) + InStr(1, fileimportacion, "."))
'            sZone = "Copy:"
'            pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
'            pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
'
'            Do While fs.FileExists(pathArchivo)
'
'                sFile.CopyFile pathArchivo, pathArchivoNuevo
'                sZone = "Delete:"
'                sFile.DeleteFile pathArchivo
'                txtCorrida.Text = vgCORRIDA
'                If Not IsNull(rsCamp("UltimaCorridaCantidadImportada")) Then txtRegistros.Text = rsCamp("UltimaCorridaCantidadImportada")
'                txtRegistros.Text = lCantidadRegistros
'                TxtActivos.Text = lCantidadActivos
'
'                Select Case nroArchivo
'                    Case 1
'                        sArchivo = "viaj"
'                    Case 2
'                        sArchivo = "COMBINADO"
'                    Case 3
'                        sArchivo = "INTEGRAL"
'                    Case 4
'                        sArchivo = "AUTO_A"
'                    Case 5
'                        sArchivo = "AUTO_B"
'                    Case 6
'                        sArchivo = "AUTO_C"
'                    Case 7
'                        sArchivo = "AUTO_E"
'                    Case 8
'                        sArchivo = "CAMIONES"
'                    Case Else
'                        Exit Do
'                End Select
'
'                pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & sExtencion
'                pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
'                nroArchivo = nroArchivo + 1
'            Loop
            
            vgCORRIDA = 0
        End If
        
        
        If vgCORRIDA > 0 Then
            sArchivo = Mid(fileimportacion, 1, Len(fileimportacion) - 4)
            'sExtencion = Mid(FileImportacion, Len(FileImportacion) - 3, Len(FileImportacion))
            sExtencion = Mid(fileimportacion, InStr(1, fileimportacion, "."), Len(fileimportacion) + InStr(1, fileimportacion, "."))
            sZone = "Copy:"
            pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & fileimportacion
            pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
            
            Do While fs.FileExists(pathArchivo)
            
'                If nroArchivo = 0 Then
'                    pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & sExtencion
'                    pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
'                Else
'                    pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & nroArchivo & sExtencion
'                    pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & nroArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
'                End If
                
                sFile.CopyFile pathArchivo, pathArchivoNuevo
                sZone = "Delete:"
                sFile.DeleteFile pathArchivo
                txtCorrida.Text = vgCORRIDA
                If Not IsNull(rsCamp("UltimaCorridaCantidadImportada")) Then txtRegistros.Text = rsCamp("UltimaCorridaCantidadImportada")
                txtRegistros.Text = lCantidadRegistros
                TxtActivos.Text = lCantidadActivos
                
                pathArchivo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & nroArchivo & sExtencion
                pathArchivoNuevo = App.Path & vgPosicionRelativa & sDirImportacion & "\" & sArchivo & nroArchivo & "_" & vgCORRIDA & "_" & sAno & sMes & sDia & sExtencion
                nroArchivo = nroArchivo + 1
            Loop
            
        End If
'==================================================
'sExtencion = Mid(FileImportacion, InStr(1, FileImportacion, ".") + Len(FileImportacion), Len(FileImportacion))
''==================================================

'    If rsCorrida.State <> 0 Then rsCorrida.Close
'    End If
'    Me.Width = 6240
'    txtProcesando.Width = 6015
'    txtProcesando.Visible = False
    sZone = ""
    Timer1.Enabled = False
    Exit Sub
errores:
        If Err.Number = 91 Then
            MsgBox "No cambió el nombre del archivo: " & Err.Description
        ElseIf sZone = "Copy:" Then
            MsgBox "No se pudo copiar el archivo original: " & Err.Description
        ElseIf sZone = "Delete:" Then
            MsgBox "No se pudo eliminar el archivo original luego de copiarlo: " & Err.Description
        Else
            MsgBox sZone & " Error de proceso: " & Err.Description
        End If
'        Me.Width = 6240
'        txtProcesando.Width = 6015
'        txtProcesando.Visible = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

s = s + 1
labelS = s
If s = 60 Then
m = m + 1
labelM = m
s = 0
End If
If m = 60 Then
h = h + 1
labelH = h
m = 0
End If

End Sub

Private Sub Form_Load()

Dim gsServidor As String, gsBaseEmpresa As String, gsBaseEmpresa1 As String, gsServidor2 As String, gsBaseEmpresa2 As String
Dim rsc As New Recordset, i As Integer
Dim ssql As String
    gsServidor = "skinner" '"(Local)" '
    gsBaseEmpresa = "BandejadeEntrada"
    gsBaseEmpresa1 = "AuxilioUT"
    gsServidor2 = "Marge"
    gsBaseEmpresa2 = "Universalt"
    
'    With txtProcesando
'        .Left = 0
'        .Top = 0
'        .Width = ScaleWidth
'        .Height = ScaleHeight
'    End With
'    cn.CommandTimeout = 36000
'    cn.ConnectionTimeout = 36000
'    cn1.CommandTimeout = 36000
'    cn1.ConnectionTimeout = 36000
'    cn2.CommandTimeout = 36000
'    cn2.ConnectionTimeout = 36000
    
    With cn
        .Provider = "sqloledb"
        .Open "Data Source = " & gsServidor & ";Initial Catalog=" & gsBaseEmpresa & ";", "sa", "cer822" '"" '
        .CommandTimeout = 0
    End With
    With cn1
        .Provider = "sqloledb"
        .Open "Data Source = " & gsServidor & ";Initial Catalog=" & gsBaseEmpresa1 & ";", "sa", "cer822" '"" '
        .CommandTimeout = 0
    End With
    
     With cn2
        .Provider = "sqloledb"
        .Open "Data Source = " & gsServidor2 & ";Initial Catalog=" & gsBaseEmpresa2 & ";", "usersponsoring", "cer822cer" '"" '
        .CommandTimeout = 0
    End With
    
    ssql = "Select distinct a.idcontacto,a.NombreYApellido from tm_Contactos a  where a.idtipodecontacto = 1 and a.activo=1 order by a.NombreYApellido"
    rsc.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    
    Do Until rsc.EOF
        cmbCia.AddItem rsc("NombreYApellido")
        rsc.MoveNext
    Loop
    vgPosicionRelativa = ""

End Sub

Private Sub TxtCoberturas_LostFocus()
TxtCoberturas.Visible = False
TxtCoberturas.Text = ""
End Sub


Private Sub txtprocesando_DblClick()
    Dim rs As New Recordset
    Dim ssql As String
    Dim respuesta As String
    
    ssql = "SELECT RegistrosLeidos "
    ssql = ssql & " FROM tm_ImportacionHistorial "
    ssql = ssql & " WHERE corrida = " & vgCORRIDA
    
    rs.Open ssql, cn1, adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then
    
        If lidCampana = 726 Then ' para Triunfo Seguros AP
            respuesta = "Se procesaron registros de Triunfo Seguros y se procesaron " & rs("RegistrosLeidos") & " registros de Triunfo Seguros AP." & Chr(13) & Chr(13) & "Saludos."
        Else
            'respuesta = "Se procesaron XXX registros." & Chr(13) & Chr(13) & "Saludos."
            respuesta = "Se procesaron " & rs("RegistrosLeidos") & " registros." & Chr(13) & Chr(13) & "Saludos."
        End If
    
        'MsgBox respuesta, vbOKOnly, "Easter egg"
        
        TxtCoberturas.Text = respuesta
        TxtCoberturas.Font = "Calibri"
        TxtCoberturas.FontSize = 11
        TxtCoberturas.Width = 3200
        TxtCoberturas.Visible = True
        
        ImportadordePolizas.Width = 8200

    End If
    
    rs.Close
    
End Sub
