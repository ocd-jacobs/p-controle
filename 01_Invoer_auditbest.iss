Sub Main
	' Neem het huidige database object
	DIM db as Object
	DIM task as Object
	Dim field As Object
	Dim table As Object
	DIM eqn as String
	Dim dbName As String

	strJaar = arg1
	strMaand = arg2
	
	' Access Import
	Const SCAN_ALL = -1
	Const SCAN_NONE = 0
	
	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	
	Set task = Client.GetImportTask("Access")
	task.InputFileName = "F:\" & strJaar & "\P-" & strJaar & "\00_Input\" & strMaand & "\P_Conversie.mdb"
	task.OutputFileNamePrefix = "Tot"
	task.CreateRecordNumberField = False
	task.DetermineMaximumCharacterFieldLengths = SCAN_ALL
	task.AddTable("DienstKantorenBoek")
	task.AddTable("tblQryStamgegevens")
	task.AddTable("Aanvulling")
	task.PerformTask
	
	Client.RunAtServer False
	Set db = Client.OpenDatabase ("F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Tot-tblQryStamgegevens.IMD")
	' Sluit huidige database
	db.Close	
	Client.RefreshFileExplorer 
End Sub
