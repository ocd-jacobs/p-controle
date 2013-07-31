Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ReportReaderImport(strJaar, strMaand)	'F:\2013\P-2013\01_Jan\onverdicht.txt
	Call FieldManipulationAppendFields()	'onverdicht.IMD
	Client.RefreshFileExplorer 
End Sub

Function ReportReaderImport(strJaar, strMaand)
	dbName = "onverdicht.IMD"
	Client.ImportPrintReport "F:\" & strJaar & "\P-" & strJaar & "\" &"onverdicht.jpm", Client.workingDirectory & "onverdicht.csv", dbname, FALSE
	Client.OpenDatabase (dbName)
End Function

Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("onverdicht.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "@if(@right(BETRG_T;1) <> ""-""; @Val(@replace(BETRG_T;""."";"",""));  @Val(@replace(@replace(BETRG_T;""."";"","");""-"";"""")) * -1)"
	field.Name = "BETRG"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = eqn
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function

