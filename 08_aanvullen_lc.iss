Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call FieldManipulationModifyRemoveFields()	'LNCMPNT-Export Worksheet.IMD
	Call JoinDatabase()	'Process-Onverdicht_KP.IMD
	Client.RefreshFileExplorer 
End Sub

Function FieldManipulationModifyRemoveFields
	Set db = Client.OpenDatabase("LNCMPNT-Export Worksheet.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	field.Name = "LC_CODE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 4
	task.ReplaceField "CODE", field
	Set field = table.NewField
	field.Name = "LC_BETEKENIS"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 18
	task.ReplaceField "BETEKENIS", field
	Set field = table.NewField
	field.Name = "LC_GELDIG_VANAF"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 8
	task.ReplaceField "GELDIG_VANAF", field
	Set field = table.NewField
	field.Name = "LC_GELDIG_TOT"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 8
	task.ReplaceField "GELDIG_TOT", field
	Set field = table.NewField
	field.Name = "LC_ACTIEF"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 1
	task.ReplaceField "ACTIEF", field
	Set field = table.NewField
	field.Name = "LC_LOONCOMPONENT"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 4
	task.ReplaceField "LOONCOMPONENT", field
	Set field = table.NewField
	field.Name = "LC_GROOTBOEKREKENING"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 6
	task.ReplaceField "GROOTBOEKREKENING", field
	Set field = table.NewField
	field.Name = "LC_GROOTBOEKREKENING_2_CENTRAAL"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 6
	task.ReplaceField "GROOTBOEKREKENING_2_CENTRAAL", field
	Set field = table.NewField
	field.Name = "LC_GROOTBOEKREK_3_DECENTRAAL_BLS"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 6
	task.ReplaceField "GROOTBOEKREK_3_DECENTRAAL_BLS", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function

Function JoinDatabase
	Set db = Client.OpenDatabase("Process-Onverdicht_KP.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "LNCMPNT-Export Worksheet.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "LGART", "LC_LOONCOMPONENT", "A"
	dbName = "Onverdicht_KP_LC.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
