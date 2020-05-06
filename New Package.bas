'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: E:\Asistencia al Vehiculo_Hogar\ImportacionDePolizas\ImportFile\New Package.bas
'Package Name: New Package
'Package Description: DTS package description
'Generated Date: 7/8/2005
'Generated Time: 5:43:35 PM
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2
Private Sub Main()
	set goPackage = goPackageOld

	goPackage.Name = "New Package"
	goPackage.Description = "DTS package description"
	goPackage.WriteCompletionStatusToNTEventLog = False
	goPackage.FailOnError = False
	goPackage.PackagePriorityClass = 2
	goPackage.MaxConcurrentSteps = 4
	goPackage.LineageOptions = 0
	goPackage.UseTransaction = True
	goPackage.TransactionIsolationLevel = 4096
	goPackage.AutoCommitTransaction = True
	goPackage.RepositoryMetadataOptions = 0
	goPackage.UseOLEDBServiceComponents = True
	goPackage.LogToSQLServer = False
	goPackage.LogServerFlags = 0
	goPackage.FailPackageOnLogFailure = False
	goPackage.ExplicitGlobalVariables = False
	goPackage.PackageType = 0
	

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection as DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

	oConnection.ConnectionProperties("Data Source") = "D:\Asistencia al Vehiculo_Hogar\ImportacionDePolizas\El Comercio\El Comercio.TXT"
	oConnection.ConnectionProperties("Mode") = 1
	oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
	oConnection.ConnectionProperties("File Format") = 2
	oConnection.ConnectionProperties("Column Lengths") = "3,13,3,50,55,30,23,81,18,55,6"
	oConnection.ConnectionProperties("File Type") = 1
	oConnection.ConnectionProperties("Skip Rows") = 0
	oConnection.ConnectionProperties("First Row Column Name") = False
	oConnection.ConnectionProperties("Number of Column") = 11
	
	oConnection.Name = "Connection 1"
	oConnection.ID = 1
	oConnection.Reusable = True
	oConnection.ConnectImmediate = False
	oConnection.DataSource = "D:\Asistencia al Vehiculo_Hogar\ImportacionDePolizas\El Comercio\El Comercio.TXT"
	oConnection.ConnectionTimeout = 60
	oConnection.UseTrustedConnection = False
	oConnection.UseDSL = False
	
	'If you have a password for this connection, please uncomment and add your password below.
	'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

	oConnection.ConnectionProperties("Persist Security Info") = True
	oConnection.ConnectionProperties("User ID") = "sa"
	oConnection.ConnectionProperties("Initial Catalog") = "AuxilioCopia"
	oConnection.ConnectionProperties("Data Source") = "OCTANS"
	oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
	
	oConnection.Name = "Connection 2"
	oConnection.ID = 2
	oConnection.Reusable = True
	oConnection.ConnectImmediate = False
	oConnection.DataSource = "OCTANS"
	oConnection.UserID = "sa"
	oConnection.ConnectionTimeout = 60
	oConnection.Catalog = "AuxilioCopia"
	oConnection.UseTrustedConnection = False
	oConnection.UseDSL = False
	
	'If you have a password for this connection, please uncomment and add your password below.
	'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep as DTS.Step2
Dim oPrecConstraint as DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

	oStep.Name = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Step"
	oStep.Description = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Step"
	oStep.ExecutionStatus = 1
	oStep.TaskName = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oStep.CommitSuccess = False
	oStep.RollbackFailure = False
	oStep.ScriptLanguage = "VBScript"
	oStep.AddGlobalVariables = True
	oStep.RelativePriority = 3
	oStep.CloseConnection = False
	oStep.ExecuteInMainThread = False
	oStep.IsPackageDSORowset = False
	oStep.JoinTransactionIfPresent = False
	oStep.DisableStep = False
	oStep.FailPackageOnError = False
	
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

	oStep.Name = "Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Step"
	oStep.Description = "Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Step"
	oStep.ExecutionStatus = 1
	oStep.TaskName = "Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oStep.CommitSuccess = False
	oStep.RollbackFailure = False
	oStep.ScriptLanguage = "VBScript"
	oStep.AddGlobalVariables = True
	oStep.RelativePriority = 3
	oStep.CloseConnection = False
	oStep.ExecuteInMainThread = False
	oStep.IsPackageDSORowset = False
	oStep.JoinTransactionIfPresent = False
	oStep.DisableStep = False
	oStep.FailPackageOnError = False
	
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Step")
	oPrecConstraint.StepName = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Step"
	oPrecConstraint.PrecedenceBasis = 1
	oPrecConstraint.Value = 0
	
oStep.precedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task (Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task)
Call Task_Sub1( goPackage	)

'------------- call Task_Sub2 for task Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task (Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task)
Call Task_Sub2( goPackage	)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
goPackage.Uninitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
set goPackage = Nothing

set goPackageOld = Nothing

End Sub


'------------- define Task_Sub1 for task Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task (Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

	oCustomTask1.Name = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oCustomTask1.Description = "Delete from Table [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oCustomTask1.SQLStatement = "delete from [AuxilioCopia].[dbo].[ImportaDatos]"
	oCustomTask1.ConnectionID = 2
	oCustomTask1.CommandTimeout = 0
	oCustomTask1.OutputAsRecordset = False
	
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task (Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

	oCustomTask2.Name = "Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oCustomTask2.Description = "Copy Data from El Comercio to [AuxilioCopia].[dbo].[ImportaDatos] Task"
	oCustomTask2.SourceConnectionID = 1
	oCustomTask2.SourceObjectName = "D:\Asistencia al Vehiculo_Hogar\ImportacionDePolizas\El Comercio\El Comercio.TXT"
	oCustomTask2.DestinationConnectionID = 2
	oCustomTask2.DestinationObjectName = "[AuxilioCopia].[dbo].[ImportaDatos]"
	oCustomTask2.ProgressRowCount = 1000
	oCustomTask2.MaximumErrorCount = 0
	oCustomTask2.FetchBufferSize = 1
	oCustomTask2.UseFastLoad = True
	oCustomTask2.InsertCommitSize = 0
	oCustomTask2.ExceptionFileColumnDelimiter = "|"
	oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
	oCustomTask2.AllowIdentityInserts = False
	oCustomTask2.FirstRow = 0
	oCustomTask2.LastRow = 0
	oCustomTask2.FastLoadOptions = 2
	oCustomTask2.ExceptionFileOptions = 1
	oCustomTask2.DataPumpOptions = 0
	
Call oCustomTask2_Trans_Sub1( oCustomTask2	)
		
		
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

	Dim oTransformation As DTS.Transformation2
	Dim oTransProps as DTS.Properties
	Dim oColumn As DTS.Column
	Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
		oTransformation.Name = "DirectCopyXform"
		oTransformation.TransformFlags = 63
		oTransformation.ForceSourceBlobsBuffered = 0
		oTransformation.ForceBlobsInMemory = False
		oTransformation.InMemoryBlobSize = 1048576
		oTransformation.TransformPhases = 4
		
		Set oColumn = oTransformation.SourceColumns.New("Col001" , 1)
			oColumn.Name = "Col001"
			oColumn.Ordinal = 1
			oColumn.Flags = 48
			oColumn.Size = 3
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col002" , 2)
			oColumn.Name = "Col002"
			oColumn.Ordinal = 2
			oColumn.Flags = 48
			oColumn.Size = 13
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col003" , 3)
			oColumn.Name = "Col003"
			oColumn.Ordinal = 3
			oColumn.Flags = 48
			oColumn.Size = 3
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col004" , 4)
			oColumn.Name = "Col004"
			oColumn.Ordinal = 4
			oColumn.Flags = 48
			oColumn.Size = 50
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col005" , 5)
			oColumn.Name = "Col005"
			oColumn.Ordinal = 5
			oColumn.Flags = 48
			oColumn.Size = 55
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col006" , 6)
			oColumn.Name = "Col006"
			oColumn.Ordinal = 6
			oColumn.Flags = 48
			oColumn.Size = 30
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col007" , 7)
			oColumn.Name = "Col007"
			oColumn.Ordinal = 7
			oColumn.Flags = 48
			oColumn.Size = 23
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col008" , 8)
			oColumn.Name = "Col008"
			oColumn.Ordinal = 8
			oColumn.Flags = 48
			oColumn.Size = 81
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col009" , 9)
			oColumn.Name = "Col009"
			oColumn.Ordinal = 9
			oColumn.Flags = 48
			oColumn.Size = 18
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col010" , 10)
			oColumn.Name = "Col010"
			oColumn.Ordinal = 10
			oColumn.Flags = 48
			oColumn.Size = 55
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.SourceColumns.New("Col011" , 11)
			oColumn.Name = "Col011"
			oColumn.Ordinal = 11
			oColumn.Flags = 48
			oColumn.Size = 6
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.SourceColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("IDPOLIZA" , 1)
			oColumn.Name = "IDPOLIZA"
			oColumn.Ordinal = 1
			oColumn.Flags = 120
			oColumn.Size = 0
			oColumn.DataType = 3
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("IDCIA" , 2)
			oColumn.Name = "IDCIA"
			oColumn.Ordinal = 2
			oColumn.Flags = 120
			oColumn.Size = 0
			oColumn.DataType = 3
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("NUMEROCOMPANIA" , 3)
			oColumn.Name = "NUMEROCOMPANIA"
			oColumn.Ordinal = 3
			oColumn.Flags = 104
			oColumn.Size = 3
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("NROPOLIZA" , 4)
			oColumn.Name = "NROPOLIZA"
			oColumn.Ordinal = 4
			oColumn.Flags = 104
			oColumn.Size = 20
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("NROSECUENCIAL" , 5)
			oColumn.Name = "NROSECUENCIAL"
			oColumn.Ordinal = 5
			oColumn.Flags = 104
			oColumn.Size = 3
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("APELLIDOYNOMBRE" , 6)
			oColumn.Name = "APELLIDOYNOMBRE"
			oColumn.Ordinal = 6
			oColumn.Flags = 104
			oColumn.Size = 255
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("DOMICILIO" , 7)
			oColumn.Name = "DOMICILIO"
			oColumn.Ordinal = 7
			oColumn.Flags = 104
			oColumn.Size = 255
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("LOCALIDAD" , 8)
			oColumn.Name = "LOCALIDAD"
			oColumn.Ordinal = 8
			oColumn.Flags = 104
			oColumn.Size = 100
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("PROVINCIA" , 9)
			oColumn.Name = "PROVINCIA"
			oColumn.Ordinal = 9
			oColumn.Flags = 104
			oColumn.Size = 100
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("CODIGOPOSTAL" , 10)
			oColumn.Name = "CODIGOPOSTAL"
			oColumn.Ordinal = 10
			oColumn.Flags = 104
			oColumn.Size = 10
			oColumn.DataType = 129
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

		Set oColumn = oTransformation.DestinationColumns.New("FECHAVIGENCIA" , 11)
			oColumn.Name = "FECHAVIGENCIA"
			oColumn.Ordinal = 11
			oColumn.Flags = 120
			oColumn.Size = 0
			oColumn.DataType = 135
			oColumn.Precision = 0
			oColumn.NumericScale = 0
			oColumn.Nullable = True
			
		oTransformation.DestinationColumns.Add oColumn
		Set oColumn = Nothing

	Set oTransProps = oTransformation.TransformServerProperties

		
	Set oTransProps = Nothing

	oCustomTask2.Transformations.Add oTransformation
	Set oTransformation = Nothing

End Sub

