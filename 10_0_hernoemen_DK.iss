Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call FieldManipulationModifyRemoveFields()	'Tot-DienstKantorenBoek.IMD
	Client.RefreshFileExplorer 
End Sub


' Gegevens: Velden Bewerken - Velden Wijzigen/Verwijderen
Function FieldManipulationModifyRemoveFields
	Set db = Client.OpenDatabase("Tot-DienstKantorenBoek.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	field.Name = "DK_A"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 2
	task.ReplaceField "A", field
	Set field = table.NewField
	field.Name = "DK_B"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 2
	task.ReplaceField "B", field
	Set field = table.NewField
	field.Name = "DK_C"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 1
	task.ReplaceField "C", field
	Set field = table.NewField
	field.Name = "DK_D"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 1
	task.ReplaceField "D", field
	Set field = table.NewField
	field.Name = "DK_E"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 1
	task.ReplaceField "E", field
	Set field = table.NewField
	field.Name = "DK_FEITALG"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 4
	task.ReplaceField "FEITALG", field
	Set field = table.NewField
	field.Name = "DK_OMSCHRIJVING"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 75
	task.ReplaceField "OMSCHRIJVING", field
	Set field = table.NewField
	field.Name = "DK_HFD"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 2
	task.ReplaceField "HFD", field
	Set field = table.NewField
	field.Name = "DK_SUB"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 2
	task.ReplaceField "SUB", field
	Set field = table.NewField
	field.Name = "DK_HOOFDSTUK"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 33
	task.ReplaceField "HOOFDSTUK", field
	Set field = table.NewField
	field.Name = "DK_DIRECTIE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 6
	task.ReplaceField "DIRECTIE", field
	Set field = table.NewField
	field.Name = "DK_DIENST"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 6
	task.ReplaceField "DIENST", field
	Set field = table.NewField
	field.Name = "DK_ARTIKEL"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Length = 2
	task.ReplaceField "ARTIKEL", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function