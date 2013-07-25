Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call JoinDatabase()	'Onverdicht_KP_LC.IMD
	Client.RefreshFileExplorer 
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Onverdicht_KP_LC.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Tot-Aanvulling.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "KOSTENPL"
	task.AddSFieldToInc "FEITALG"
	task.AddMatchKey "PERNR", "PERSNR", "A"
	dbName = "Onverdicht_KP_LC_Aanv.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function