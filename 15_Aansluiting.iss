Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call Summarization(strMaand)	'Onverdicht_KP_LC_Aanv.IMD
	Call ExportDatabaseXLS8(strJaar, strMaand)	'Aansluiting P-Jurist 01_Jan.IMD
	Client.RefreshFileExplorer 
End Sub


' Analyse: Totaliseren
Function Summarization(strMaand)
	Set db = Client.OpenDatabase("Onverdicht_KP_LC_Aanv.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "DIENSTKANTOOR_LANG"
	task.AddFieldToSummarize "KOSTENSOORT_NUM"
	task.AddFieldToTotal "BETRG"
	dbName = "Aansluiting P-Jurist " & strMaand & ".IMD"
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
	Set db = Client.OpenDatabase("Aansluiting P-Jurist " & strMaand & ".IMD")
	Set task = db.ExportDatabase
	task.AddFieldToInc "KOSTENSOORT_NUM"
	task.AddFieldToInc "DIENSTKANTOOR_LANG"
	task.AddFieldToInc "BETRG_SOM"
	eqn = ""
	task.PerformTask "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Aansluiting P-Jurist " & strMaand & ".XLS", "Database", "XLS8", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function