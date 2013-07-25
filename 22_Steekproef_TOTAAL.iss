Sub Main
	Call AppendDatabase()	'MUS massa 2 object 6.IMD
	Client.CloseDatabase "MUS.IMD"
End Sub


' Bestand: Databases Samenvoegen
Function AppendDatabase
	Set db = Client.OpenDatabase("MUS massa 1 object 1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "MUS massa 1 object 2.IMD"
	task.AddDatabase "MUS massa 1 object 3.IMD"
	task.AddDatabase "MUS massa 1 object 4.IMD"
	task.AddDatabase "MUS massa 1 object 5.IMD"
	task.AddDatabase "MUS massa 1 object 6 Art 29.IMD"
	task.AddDatabase "MUS massa 1 object 6 Art 91.IMD"
	task.AddDatabase "MUS massa 2 object 1.IMD"
	task.AddDatabase "MUS massa 2 object 2.IMD"
	task.AddDatabase "MUS massa 2 object 3.IMD"
	task.AddDatabase "MUS massa 2 object 4.IMD"
	task.AddDatabase "MUS massa 2 object 5.IMD"
	task.AddDatabase "MUS massa 2 object 6 Art 29.IMD"
	task.AddDatabase "MUS massa 2 object 6 Art 91.IMD"
	dbName = "MUS.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function