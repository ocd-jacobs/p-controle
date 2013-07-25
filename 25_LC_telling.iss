Sub Main
	Call Summarization()	'Onverdicht.IMD
End Sub


' Analyse: Totaliseren
Function Summarization
	Set db = Client.OpenDatabase("Onverdicht.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "LGART"
	dbName = "LC telling.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_COUNT
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function