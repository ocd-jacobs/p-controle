Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call JoinDatabase()	'Process-Onverdicht_KP.IMD
	Client.RefreshFileExplorer 
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Process-Onverdicht_KP.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "LC.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "LGART", "LC_LGART", "A"
	dbName = "Onverdicht_KP_LC.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function