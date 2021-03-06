Sub Main
	Call MUSExtraction1()	'Massa=1 + 6.IMD
	Call MUSExtraction2()	'Massa=1 + 6.IMD
End Sub


' Steekproef: Geldeenheid
Function MUSExtraction1
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_RangeOfValues_POSITIVES = 0
	Const WI_RangeOfValues_NEGATIVES = 1
	Const WI_RangeOfValues_ABSOLUTES = 2
	Const WI_TaskType_FIXED = 0
	Const WI_TaskType_CELL = 1
	
	Set db = Client.OpenDatabase("Massa=1 + 6 Art 29.IMD")
	Set task = db.MUSExtraction
	task.IncludeAllFields
	task.TaskType = WI_TaskType_CELL
	task.RangeOfValues = WI_RangeOfValues_ABSOLUTES
	task.HighValueHandling = WI_HighValueHandling_AGGREGATE
	
	DAD_Interval = 110000
	DAD_seed = Round(Rnd() * 0.5 * DAD_Interval, 0)
	
	task.SampleInterval = DAD_Interval
	task.RandomValue = DAD_seed
	task.FieldToSample = "BETRG"
	dbName = "MUS massa 1 object 6 Art 29.IMD"
	task.MUSExtractionFilename = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Steekproef: Geldeenheid
Function MUSExtraction2
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_RangeOfValues_POSITIVES = 0
	Const WI_RangeOfValues_NEGATIVES = 1
	Const WI_RangeOfValues_ABSOLUTES = 2
	Const WI_TaskType_FIXED = 0
	Const WI_TaskType_CELL = 1
	
	Set db = Client.OpenDatabase("Massa=1 + 6 Art 91.IMD")
	Set task = db.MUSExtraction
	task.IncludeAllFields
	task.TaskType = WI_TaskType_CELL
	task.RangeOfValues = WI_RangeOfValues_ABSOLUTES
	task.HighValueHandling = WI_HighValueHandling_AGGREGATE
	
	DAD_Interval = 3380000.00
	DAD_seed = Round(Rnd() * DAD_Interval, 0)
	
	task.SampleInterval = DAD_Interval
	task.RandomValue = DAD_seed

	task.FieldToSample = "BETRG"
	dbName = "MUS massa 1 object 6 Art 91.IMD"
	task.MUSExtractionFilename = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
