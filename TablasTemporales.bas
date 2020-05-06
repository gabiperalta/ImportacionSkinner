Attribute VB_Name = "TablasTemporales"
Public Function creaTabla(vdescCampana As String) As Boolean
On Error Resume Next

creaTabla = True

vdescCampana = Replace(vdescCampana, " ", "")

cn.Execute "drop table ImportaDatos" & vdescCampana

If Err.Number Then Err.Clear

NewTable = "CREATE TABLE [dbo].[ImportaDatos" & vdescCampana & "] (                         "
NewTable = NewTable & "    [IDPOLIZA] [int] NULL ,                                                               "
NewTable = NewTable & "    [IDCIA] [int] NULL ,                                                                  "
NewTable = NewTable & "    [NUMEROCOMPANIA] [varchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,                    "
NewTable = NewTable & "    [NROPOLIZA] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,                        "
NewTable = NewTable & "    [NROSECUENCIAL] [varchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,                     "
NewTable = NewTable & "    [APELLIDOYNOMBRE] [varchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,                 "
NewTable = NewTable & "    [FECHADENACIMIENTO] [datetime] NULL ,                                                  "
NewTable = NewTable & "    [DOMICILIO] [varchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [LOCALIDAD] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [PROVINCIA] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [CODIGOPOSTAL] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,                     "
NewTable = NewTable & "    [FECHAVIGENCIA] [datetime] NULL ,                                                     "
NewTable = NewTable & "    [FECHAVENCIMIENTO] [datetime] NULL ,                                                  "
NewTable = NewTable & "    [FECHAALTAOMNIA] [datetime] NULL ,                                                    "
NewTable = NewTable & "    [FECHABAJAOMNIA] [datetime] NULL ,                                                    "
NewTable = NewTable & "    [IDAUTO] [int] NULL ,                                                                 "
NewTable = NewTable & "    [MARCADEVEHICULO] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,                "
NewTable = NewTable & "    [MODELO] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                          "
NewTable = NewTable & "    [COLOR] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,                            "
NewTable = NewTable & "    [ANO] [varchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,                               "
NewTable = NewTable & "    [PATENTE] [varchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,                          "
NewTable = NewTable & "    [TIPODEVEHICULO] [int] NULL ,                                                         "
NewTable = NewTable & "    [TipodeServicio] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [IDTIPODECOBERTURA] [int] NULL ,                                                      "
NewTable = NewTable & "    [COBERTURAVEHICULO] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,                    "
NewTable = NewTable & "    [COBERTURAVIAJERO] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,                     "
NewTable = NewTable & "    [TipodeOperacion] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,                     "
NewTable = NewTable & "    [Operacion] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,                           "
NewTable = NewTable & "    [CATEGORIA] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,                         "
NewTable = NewTable & "    [ASISTENCIAXENFERMEDAD] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,                "
NewTable = NewTable & "    [CORRIDA] [int] NULL ,                                                                "
NewTable = NewTable & "    [FECHACORRIDA] [datetime] NULL ,                                                      "
NewTable = NewTable & "    [IdCampana] [int] NULL ,                                                              "
NewTable = NewTable & "    [Conductor] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                        "
NewTable = NewTable & "    [CodigoDeProductor] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,                "
NewTable = NewTable & "    [CodigoDeServicioVip] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,               "
NewTable = NewTable & "    [TipodeDocumento] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [NumeroDeDocumento] [varchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,                "
NewTable = NewTable & "    [TipodeHogar] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,                          "
NewTable = NewTable & "    [IniciodeAnualidad] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,                   "
NewTable = NewTable & "    [PolizaIniciaAnualidad] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,               "
NewTable = NewTable & "    [Telefono] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,                            "
NewTable = NewTable & "    [NroMotor] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                        "
NewTable = NewTable & "    [Gama] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                            "
NewTable = NewTable & "    [IdLote] [int] NULL ,                                                                 "
NewTable = NewTable & "    [InformadoSinCobertura] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,            "
NewTable = NewTable & "    [MontoCoverturaVidrios] [float] NULL ,                                                "
NewTable = NewTable & "    [COBERTURAHOGAR] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,                   "
NewTable = NewTable & "    [CodigoDeProceso] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [IdTipodePoliza] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [Referido] [int] NULL ,                                                               "
NewTable = NewTable & "    [Telefono2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,                           "
NewTable = NewTable & "    [Telefono3] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,                           "
NewTable = NewTable & "    [IdProducto] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,                       "
NewTable = NewTable & "    [email] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                            "
NewTable = NewTable & "    [email2] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                           "
NewTable = NewTable & "    [pais] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,                             "
NewTable = NewTable & "    [FechaNac] [datetime] NULL ,                                                          "
NewTable = NewTable & "    [Modificaciones] [int] NULL ,                                                         "
NewTable = NewTable & "    [Importe] [float] NULL ,                                                "
NewTable = NewTable & "    [Sexo] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,                                 "
NewTable = NewTable & "    [Agencia] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                          "
NewTable = NewTable & "    [Codigoencliente] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [Ocupacion] [varchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [Certificado] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [NroSecunecialEnCliente] [varchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,                  "
NewTable = NewTable & "    [DocumentoReferente] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL                 "
NewTable = NewTable & ") ON [PRIMARY]                                                                            "

cn.Execute NewTable

If Err.Number Then
    creaTabla = False
End If

End Function
