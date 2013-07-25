Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call AccessImport(strJaar, strMaand)	'F:\2009\P-2009\12_Dec\Process_Onverdicht.MDB
	Client.RefreshFileExplorer 
End Sub


' Bestand - Import Assistent: Access
Function AccessImport(strJaar, strMaand)
	Const SCAN_ALL = -1
	Const SCAN_NONE = 0
	
	Set task = Client.GetImportTask("Access")
	task.InputFileName = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Process_Onverdicht.MDB"
	task.OutputFileNamePrefix = "Process"
	task.CreateRecordNumberField = False
	task.DetermineMaximumCharacterFieldLengths = 10000
	task.AddTable("Onverdicht_KP")
	task.PerformTask
	dbName = task.OutputFileNameFromTableName("Onverdicht_KP")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function