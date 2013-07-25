Sub Main
	Call KeyValueExtraction()	'Onverdicht_Object_Aanv.IMD
End Sub


' Gegevens: Sleutelwaarde Selectie
Function KeyValueExtraction
	Set db = Client.OpenDatabase("Onverdicht_Object_Aanv.IMD")
	Set task = db.KeyValueExtraction
	dim myArray(23,1)
	myArray(0,0) = "0"
	myArray(0,1) = "1"
	myArray(1,0) = "0"
	myArray(1,1) = "2"
	myArray(2,0) = "0"
	myArray(2,1) = "3"
	myArray(3,0) = "0"
	myArray(3,1) = "4"
	myArray(4,0) = "0"
	myArray(4,1) = "5"
	myArray(5,0) = "0"
	myArray(5,1) = "6"
	myArray(6,0) = "1"
	myArray(6,1) = "1"
	myArray(7,0) = "1"
	myArray(7,1) = "2"
	myArray(8,0) = "1"
	myArray(8,1) = "3"
	myArray(9,0) = "1"
	myArray(9,1) = "4"
	myArray(10,0) = "1"
	myArray(10,1) = "5"
	myArray(11,0) = "1"
	myArray(11,1) = "6"
	myArray(12,0) = "2"
	myArray(12,1) = "1"
	myArray(13,0) = "2"
	myArray(13,1) = "2"
	myArray(14,0) = "2"
	myArray(14,1) = "3"
	myArray(15,0) = "2"
	myArray(15,1) = "4"
	myArray(16,0) = "2"
	myArray(16,1) = "5"
	myArray(17,0) = "2"
	myArray(17,1) = "6"
	myArray(18,0) = "3"
	myArray(18,1) = "1"
	myArray(19,0) = "3"
	myArray(19,1) = "2"
	myArray(20,0) = "3"
	myArray(20,1) = "3"
	myArray(21,0) = "3"
	myArray(21,1) = "4"
	myArray(22,0) = "3"
	myArray(22,1) = "5"
	myArray(23,0) = "3"
	myArray(23,1) = "6"
	task.IncludeAllFields
	task.AddKey "MASSA", "A"
	task.AddKey "DAD_OBJECT", "A"
	task.DBPrefix = "Massa"
	task.CreateMultipleDatabases = TRUE
	task.ValuesToExtract myArray
	task.PerformTask
	dbName = task.DBName
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
End Function