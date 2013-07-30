Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ExcelImport()	'F:\2013\P-2013\01_Jan\PKSTNPLTS.xls
	Client.RefreshFileExplorer 
End Sub

Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = Client.LocateInputFile (Client.workingDirectory & "\" & "PKSTNPLTS.xls")
	'dbName = Client.LocateInputFile ("F:\" & strJaar & "\P-" & strJaar & "\" & strMaand  & "\" & "PKSTNPLTS.xls")
	task.FileToImport = dbName
	task.SheetToImport = "Export Worksheet"
	task.OutputFilePrefix = "PKSTNPLTS"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "FALSE"
	task.PerformTask
	dbName = task.OutputFilePath("Export Worksheet")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function
