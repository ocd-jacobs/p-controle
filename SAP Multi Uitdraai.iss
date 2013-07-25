Dim strJaar As String
Dim strPost As String
Dim strNaam As String
Dim strMaand As String
Dim strPersNr As String

Sub Main
	Dim arrMonths(1 To 12) As String
	Dim intStart As Integer
	Dim intEnd As Integer
	Dim IntCount As Integer
	
	arrMonths(1) = "01_Jan"
	arrMonths(2) = "02_Feb"
	arrMonths(3) = "03_Mrt"
	arrMonths(4) = "04_Apr"
	arrMonths(5) = "05_Mei"
	arrMonths(6) = "06_Jun"
	arrMonths(7) = "07_Jul"
	arrMonths(8) = "08_Aug"
	arrMonths(9) = "09_Sep"
	arrMonths(10) = "10_Okt"
	arrMonths(11) = "11_Nov"
	arrMonths(12) = "12_Dec"
	
	'********** Begin Parameters ***********
	
	IntStart = 1
	IntEnd = 2
	
	strJaar = "2013"
		
	strPost = "X"
	strNaam = "Bakker"
	strpersNr = "20000819"

	'********** Einde Parameters ***********

	For intCount = intStart To intEnd
		strMaand = arrMonths(IntCount)
		
		Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
		'Client.workingDirectory = "F:\2011\P-2011\01_Jan" 
			
		Call JoinDatabase()	'Onverdicht.IMD
		Call ExportDatabaseXLS8()	'20002082 - jun.IMD
	Next
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Onverdicht.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Looncomponent omschr.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "OMSCHRIJVING"
	task.AddMatchKey "LGART", "LC", "A"
	task.Criteria = "PERNR = " & Chr(34) & strPersNr & Chr(34)
	dbName = "Post " & strPost & " - " & strNaam & " - " & strMaand & ".IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand - Database Exporteren: XLS8
Function ExportDatabaseXLS8
	Set db = Client.OpenDatabase("Post " & strPost & " - " & strNaam & " - " & strMaand & ".IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "H:\"& "Post " & strPost & " - " & strNaam & " - " & strMaand & ".xls", "Database", "XLS8", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function