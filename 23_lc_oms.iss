Sub Main
	Call JoinDatabase()	'Steekproef.IMD
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("MUS.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Looncomponent omschr.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "LGART", "LC", "A"
	dbName = "MUS aangevuld.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function