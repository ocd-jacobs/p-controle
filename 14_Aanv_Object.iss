Sub Main
	Call JoinDatabase()	'Onverdicht excl RVDR.IMD
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Onverdicht excl RVDR.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Objecten.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "DAD_OBJECT"
	task.AddMatchKey "DK_DIENST", "DK_DIENST", "A"
	dbName = "Onverdicht_Object.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function