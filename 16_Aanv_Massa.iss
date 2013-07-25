Sub Main
	Call JoinDatabase()	'Onverdicht_Object.IMD
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Onverdicht_Object.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "LC_Steek.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "MASSA"
	task.AddMatchKey "LGART", "LC", "A"
	dbName = "Onverdicht_Object_Aanv.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function