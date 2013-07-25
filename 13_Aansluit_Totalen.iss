Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call Summarization()	'Onverdicht excl RVDR.IMD
	Call ExportDatabaseXLS8(strJaar, strMaand)	'Aansluit totalen jan.IMD
	Client.RefreshFileExplorer 
End Sub

' Analyse: Totaliseren
Function Summarization
	Set db = Client.OpenDatabase("Onverdicht excl RVDR.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "LGART"
	task.AddFieldToSummarize "VERDICHTNGS_KS"
	task.AddFieldToSummarize "DK_DIENST"
	task.AddFieldToSummarize "DK_ARTIKEL"
	task.AddFieldToTotal "BETRG"
	dbName = "Aansluit totalen.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_COUNT + SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand - Database Exporteren: XLS8
Function ExportDatabaseXLS8(strJaar, strMaand)
	Set db = Client.OpenDatabase("Aansluit totalen.IMD")
	Set task = db.ExportDatabase
	task.AddFieldToInc "DK_DIENST"
	task.AddFieldToInc "VERDICHTNGS_KS"
	task.AddFieldToInc "LGART"
	task.AddFieldToInc "DK_ARTIKEL"
	task.AddFieldToInc "BETRG_SOM"
	eqn = ""
	task.PerformTask "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Aansluit totalen.XLS", "Database", "XLS8", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
