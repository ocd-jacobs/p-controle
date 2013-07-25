Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ExportDatabaseMDB2000(strJaar, strMaand)	'MUS jan.IMD
End Sub

' Bestand - Database Exporteren: MDB2000
Function ExportDatabaseMDB2000(strJaar, strMaand)
	Set db = Client.OpenDatabase("Steekproef.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Steekproef.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
