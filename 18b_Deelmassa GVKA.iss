Sub Main
	Call DirectExtraction1()	'Massa=1 + 6.IMD
	Call DirectExtraction2()	'Massa=1 + 6.IMD
End Sub


' Gegevens: Directe Selectie
Function DirectExtraction1
	Set db = Client.OpenDatabase("Massa=1 + 6.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Massa=1 + 6 Art 29.IMD"
	task.AddExtraction dbName, "", "DK_ARTIKEL = ""29"""
	task.AddExtraction "Massa=1 + 6 Art 91.IMD", "", "DK_ARTIKEL = ""91"""
	task.AddExtraction "Massa=1 + 6 Overig.IMD", "", "DK_ARTIKEL <> ""29"" .AND. DK_ARTIKEL <> ""91"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Gegevens: Directe Selectie
Function DirectExtraction2
	Set db = Client.OpenDatabase("Massa=2 + 6.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Massa=2 + 6 Art 29.IMD"
	task.AddExtraction dbName, "", "DK_ARTIKEL = ""29"""
	task.AddExtraction "Massa=2 + 6 Art 91.IMD", "", "DK_ARTIKEL = ""91"""
	task.AddExtraction "Massa=2 + 6 Overig.IMD", "", "DK_ARTIKEL <> ""29"" .AND. DK_ARTIKEL <> ""91"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
