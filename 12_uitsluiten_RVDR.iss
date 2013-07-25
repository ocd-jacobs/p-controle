Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call DirectExtraction()	'Onverdicht_KP_LC_DK.IMD
	Client.RefreshFileExplorer 
End Sub


' Gegevens: Directe Selectie
Function DirectExtraction
	Set db = Client.OpenDatabase("Onverdicht_KP_LC_DK.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Onverdicht excl RVDR.IMD"
	task.AddExtraction dbName, "", "@Upper(@trim(DK_DIENST)) <> ""RVDR"""
	task.AddExtraction "Onverdicht RVDR.IMD", "", "@Upper(@trim(DK_DIENST)) = ""RVDR"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function