Sub Main
	Call MUSExtraction()	'Massa=1 + 1.IMD
End Sub


' Steekproef: Geldeenheid
Function MUSExtraction
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_RangeOfValues_POSITIVES = 0
	Const WI_RangeOfValues_NEGATIVES = 1
	Const WI_RangeOfValues_ABSOLUTES = 2
	Const WI_TaskType_FIXED = 0
	Const WI_TaskType_CELL = 1
	
	Set db = Client.OpenDatabase("Massa=1 + 1.IMD")
	Set task = db.MUSExtraction
	task.IncludeAllFields
	task.TaskType = WI_TaskType_CELL
	task.RangeOfValues = WI_RangeOfValues_ABSOLUTES
	task.HighValueHandling = WI_HighValueHandling_AGGREGATE
	task.SampleInterval = 5720000.00
	task.RandomValue = 30884
	task.FieldToSample = "BETRG"
	dbName = "MUS massa 1 object 1.IMD"
	task.MUSExtractionFilename = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function