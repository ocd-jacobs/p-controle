Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ExportDatabaseMDB2000(strJaar, strMaand)	'Onverdicht.IMD
	Call ExportDatabaseMDB20001(strJaar, strMaand)	'KP.IMD
	Client.RefreshFileExplorer 
End Sub


' Bestand - Database Exporteren: MDB2000
Function ExportDatabaseMDB2000(strJaar, strMaand)
	Set db = Client.OpenDatabase("Onverdicht.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Onverdicht.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function

' Bestand - Database Exporteren: MDB2000
Function ExportDatabaseMDB20001(strJaar, strMaand)
	Set db = Client.OpenDatabase("PKSTNPLTS-Export Worksheet.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\KP.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function